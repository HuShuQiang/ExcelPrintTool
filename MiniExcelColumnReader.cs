using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using MiniExcelLibs;

namespace Excel打印辅助工具
{
    /// <summary>
    /// 用 MiniExcel 流式读取 xlsx 文件中「指定列名」那一列数据。
    ///
    /// 设计要点（针对 10w+ 数据）：
    ///   1. MiniExcel.Query 返回 IEnumerable&lt;dynamic&gt;（每行是 ExpandoObject），
    ///      底层一次性把整个 sheet 反序列化为对象，但每行只解析一次；
    ///   2. yield 出来的 IEnumerable 是惰性的，调用方消费才触发对 xlsx 的解析；
    ///   3. 调用方拿到的 IEnumerable&lt;object&gt; 可以 pipe 到 StreamWriter 写 CSV，
    ///      整个流水不会把全量 18w 行收进 List 一次性分配内存。
    ///
    /// 用法（最常见：把那一列直接落 CSV）：
    ///   int rows = MiniExcelColumnReader.ExtractColumnToCsv(
    ///       @"C:\data\boxlist.xlsx", "sn", @"C:\out\sn.csv",
    ///       progress: new Progress&lt;int&gt;(n =&gt; Console.WriteLine($"已处理 {n} 行")));
    /// </summary>
    public static class MiniExcelColumnReader
    {
        /// <summary>
        /// 流式读取 xlsx 指定列的所有值。
        /// </summary>
        /// <param name="xlsxPath">xlsx 文件路径</param>
        /// <param name="columnName">目标列名（与表头行单元格相等）</param>
        /// <param name="sheetName">可选：工作表名称；null = 第一个 Sheet</param>
        /// <param name="useHeaderRow">是否以第一行为列名（默认 true）</param>
        /// <param name="validateColumn">是否先校验列名是否存在（默认 true）。列名已由调用方校验过时传 false，
        /// 可省掉一次整表解析——分批读取 10w+ 行时这一步会被放大上千倍。</param>
        /// <returns>惰性 IEnumerable&lt;object&gt;，消费时一次解析一行</returns>
        public static IEnumerable<object> ReadColumn(
            string xlsxPath,
            string columnName,
            string sheetName = null,
            bool useHeaderRow = true,
            bool validateColumn = true)
        {
            return ReadColumnCore(xlsxPath, columnName, sheetName, useHeaderRow, validateColumn);
        }

        private static IEnumerable<object> ReadColumnCore(
            string xlsxPath,
            string columnName,
            string sheetName,
            bool useHeaderRow,
            bool validateColumn)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath))
                throw new ArgumentException("xlsx 文件路径不能为空", nameof(xlsxPath));
            if (string.IsNullOrWhiteSpace(columnName))
                throw new ArgumentException("列名不能为空", nameof(columnName));
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException("找不到 xlsx 文件", xlsxPath);

            // 1. 先确认该列存在，避免读 18w 行才发现列名错
            if (validateColumn)
                ValidateColumnExists(xlsxPath, columnName, sheetName, useHeaderRow);

            // 2. 真正的流式：每一行交给调用方后才解析
            var rows = MiniExcel.Query(xlsxPath, useHeaderRow: useHeaderRow, sheetName: sheetName);
            foreach (var row in rows)
            {
                if (row == null) continue;

                if (row is IDictionary<string, object> dict)
                {
                    if (dict.TryGetValue(columnName, out object v))
                        yield return v;
                }
                else if (row is IDictionary<string, object> dict2)
                {
                    if (dict2.TryGetValue(columnName, out object v))
                        yield return v;
                }
                // 兜底：跳过非字典行
            }
        }

        /// <summary>
        /// 列出该 xlsx 第一行（表头）的所有列名，方便调用方确认列名是否正确。
        /// </summary>
        public static IReadOnlyList<string> GetColumnNames(
            string xlsxPath, string sheetName = null, bool useHeaderRow = true)
        {
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException("找不到 xlsx 文件", xlsxPath);

            var rows = MiniExcel.Query(xlsxPath, useHeaderRow: useHeaderRow, sheetName: sheetName);
            foreach (var row in rows)
            {
                if (row is IDictionary<string, object> dict)
                    return dict.Keys.ToList();
                break;
            }
            return Array.Empty<string>();
        }

        /// <summary>
        /// 把指定列导出为 CSV（UTF-8 BOM，逗号分隔，自动双引号包裹）。
        /// 流式逐行写出，不把全量预读进内存。
        /// </summary>
        /// <returns>成功写入的行数（不含表头）</returns>
        public static int ExtractColumnToCsv(
            string xlsxPath,
            string columnName,
            string csvOutputPath,
            string sheetName = null,
            IProgress<int> progress = null,
            int progressEvery = 5000,
            bool includeHeader = true)
        {
            if (string.IsNullOrWhiteSpace(csvOutputPath))
                throw new ArgumentException("CSV 输出路径不能为空", nameof(csvOutputPath));

            string dir = Path.GetDirectoryName(csvOutputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            int n = 0;
            using (var sw = new StreamWriter(csvOutputPath, false, new UTF8Encoding(true)))
            {
                if (includeHeader)
                    sw.WriteLine(EscapeCsv(columnName));

                foreach (object v in ReadColumn(xlsxPath, columnName, sheetName))
                {
                    sw.WriteLine(EscapeCsv(v?.ToString()));
                    n++;

                    if (progress != null && (n % progressEvery == 0))
                        progress.Report(n);
                }
            }
            return n;
        }

        /// <summary>
        /// 把列数据写入 List&lt;string&gt;，适合「后续立即处理全部数据」的便捷版本。
        /// </summary>
        public static List<string> ReadColumnAsList(
            string xlsxPath, string columnName, string sheetName = null)
        {
            var list = new List<string>();
            foreach (object v in ReadColumn(xlsxPath, columnName, sheetName))
                list.Add(v?.ToString());
            return list;
        }

        // ----------------------------------------------------------------
        //  指定行数 / 范围 读取
        // ----------------------------------------------------------------

        /// <summary>
        /// 只读取指定列的前 <paramref name="count"/> 条数据（count &gt; 0）。
        /// 流式：读到第 count 条就 yield break，MiniExcel 不会构造后续行。
        /// </summary>
        /// <example>
        ///   foreach (var v in ReadColumnTop(path, "sn", 100)) { ... }
        /// </example>
        public static IEnumerable<object> ReadColumnTop(
            string xlsxPath, string columnName, int count, string sheetName = null)
        {
            if (count <= 0) yield break;
            int n = 0;
            foreach (object v in ReadColumn(xlsxPath, columnName, sheetName))
            {
                yield return v;
                if (++n >= count) yield break;
            }
        }

        /// <summary>
        /// 跳过 <paramref name="skip"/> 条之后，再读取 <paramref name="take"/> 条。
        /// skip 与 take 都基于「数据行」（不含表头）。
        /// 例：skip=10, take=5 → 数据行第 11~15 条。
        /// </summary>
        public static IEnumerable<object> ReadColumnSlice(
            string xlsxPath, string columnName, int skip, int take, string sheetName = null, bool validateColumn = true)
        {
            if (take <= 0) yield break;

            int n = 0;             // 已迭代计数
            int emitted = 0;       // 已 yield 计数
            foreach (object v in ReadColumn(xlsxPath, columnName, sheetName, true, validateColumn))
            {
                if (n >= skip && emitted < take)
                {
                    yield return v;
                    emitted++;
                    if (emitted >= take) yield break;
                }
                n++;
            }
        }

        /// <summary>
        /// 按 1-based 行号范围读取（含两端）。语义同 UI 里「第 startRow 行 ~ 第 endRow 行」。
        /// 例：startRow=11, endRow=20 → 第 11~20 条数据行。
        /// </summary>
        public static IEnumerable<object> ReadColumnRange(
            string xlsxPath, string columnName, int startRow, int endRow, string sheetName = null, bool validateColumn = true)
        {
            if (startRow < 1) startRow = 1;
            if (endRow < startRow)
                throw new ArgumentException(
                    "endRow(" + endRow + ") 必须不小于 startRow(" + startRow + ")",
                    nameof(endRow));

            int skip = startRow - 1;
            int take = endRow - startRow + 1;
            foreach (object v in ReadColumnSlice(xlsxPath, columnName, skip, take, sheetName, validateColumn))
                yield return v;
        }

        /// <summary>
        /// 同步版本：直接读取「前 count 条」并返回 List&lt;string&gt;。
        /// 适合「想马上拿到前 N 条字符串数组」的场景，比如预览。
        /// </summary>
        public static List<string> ReadColumnTopAsList(
            string xlsxPath, string columnName, int count, string sheetName = null)
        {
            var list = new List<string>(Math.Max(count, 0));
            foreach (object v in ReadColumnTop(xlsxPath, columnName, count, sheetName))
                list.Add(v?.ToString());
            return list;
        }

        /// <summary>
        /// 同步版本：跳过 <paramref name="skip"/> 条之后，再读取 <paramref name="take"/> 条，
        /// 转成字符串一次性返回。
        /// 适用场景：流式读取不方便一次性传到下游、需要 List 给 DataGridView/WPF 绑定或返回 API。
        /// </summary>
        public static List<string> ReadColumnSliceAsList(
            string xlsxPath, string columnName, int skip, int take,
            string sheetName = null, bool validateColumn = true)
        {
            if (take <= 0) return new List<string>(0);
            var list = new List<string>(take);
            foreach (object v in ReadColumnSlice(xlsxPath, columnName, skip, take, sheetName, validateColumn))
                list.Add(v?.ToString());
            return list;
        }

        /// <summary>
        /// 同步版本：按 1-based 行号范围（含两端）读取，转成字符串一次性返回。
        /// 语义同 UI 里「第 startRow 行 ~ 第 endRow 行」。
        /// </summary>
        /// <example>
        ///   var list = MiniExcelColumnReader.ReadColumnRangeAsList(path, "sn", 11, 20);
        /// </example>
        public static List<string> ReadColumnRangeAsList(
            string xlsxPath, string columnName, int startRow, int endRow,
            string sheetName = null, bool validateColumn = true)
        {
            // ReadColumnRange 自身会校验 startRow<=endRow 并负责抛出
            var list = new List<string>(Math.Max(endRow - startRow + 1, 0));
            foreach (object v in ReadColumnRange(xlsxPath, columnName, startRow, endRow, sheetName, validateColumn))
                list.Add(v?.ToString());
            return list;
        }

        /// <summary>
        /// 同步版本：读取整列并返回 List&lt;string&gt;（内部就是 <see cref="ReadColumnAsList"/>）。
        /// </summary>
        public static List<string> ReadAllAsList(
            string xlsxPath, string columnName, string sheetName = null, bool validateColumn = true)
        {
            var list = new List<string>();
            foreach (object v in ReadColumn(xlsxPath, columnName, sheetName, true, validateColumn))
                list.Add(v?.ToString());
            return list;
        }

        // ----------------------------------------------------------------
        //  helpers
        // ----------------------------------------------------------------

        /// <summary>校验列是否存在；不存在时抛出 ArgumentException，提示当前所有列名。</summary>
        private static void ValidateColumnExists(
            string xlsxPath, string columnName, string sheetName, bool useHeaderRow)
        {
            var names = GetColumnNames(xlsxPath, sheetName, useHeaderRow);
            // 大小写不敏感：excel 列名可能大小写不一
            bool exists = names.Any(k => string.Equals(k, columnName, StringComparison.OrdinalIgnoreCase));
            if (!exists)
            {
                string hint = names.Count == 0 ? "（表头为空）" : "\n  - " + string.Join("\n  - ", names);
                throw new ArgumentException(
                    "列「" + columnName + "」在文件中不存在。当前列名：" + hint
                    + "\n请检查列名（含大小写、空格）",
                    nameof(columnName));
            }
        }

        /// <summary>按 CSV 规则转义：含逗号 / 双引号 / 换行时整个字段加双引号包裹，内部双引号替换为两个双引号。</summary>
        private static string EscapeCsv(string value)
        {
            if (value == null) return "";
            bool needQuote = value.IndexOf(',') >= 0 || value.IndexOf('"') >= 0 ||
                             value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0;
            if (!needQuote) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}

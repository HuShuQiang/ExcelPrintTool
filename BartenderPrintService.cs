using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
#if BARTTENDER
using Seagull.BarTender.Print;
#endif

namespace Excel打印辅助工具
{
    /// <summary>打印进度快照（用于回传 UI）。</summary>
    public sealed class PrintProgress
    {
        /// <summary>已处理的数据行数。</summary>
        public int Rows { get; set; }

        /// <summary>已打印的标签张数（行数 × 份数）。</summary>
        public int Labels { get; set; }

        /// <summary>当前处理到第几条（相对本次打印范围的序号，从 1 开始）。</summary>
        public int CurrentIndex { get; set; }
    }

    /// <summary>批量打印的统计结果。</summary>
    public sealed class PrintBatchResult
    {
        /// <summary>成功打印的行数。</summary>
        public int PrintedRows { get; internal set; }

        /// <summary>成功打印的标签张数。</summary>
        public int PrintedLabels { get; internal set; }

        /// <summary>打印失败的行数。</summary>
        public int FailedRows { get; internal set; }

        /// <summary>是否被取消。</summary>
        public bool Canceled { get; internal set; }

        /// <summary>每一行的失败原因。</summary>
        public List<string> Errors { get; } = new List<string>();

        /// <summary>第一条错误信息。</summary>
        public string FirstError => Errors.Count > 0 ? Errors[0] : null;
    }

    /// <summary>
    /// 标签打印服务抽象：把 Bartender SDK 隔离在实现层，窗体只依赖这个接口。
    /// </summary>
    public interface ILabelPrintService : IDisposable
    {
        /// <summary>当前是否已打开模板。</summary>
        bool IsTemplateOpen { get; }

        /// <summary>当前已打开的模板路径（未打开时为 null）。</summary>
        string TemplatePath { get; }

        /// <summary>打开并常驻复用打印模板；同一路径重复调用不会重新打开。</summary>
        void OpenTemplate(string templatePath, string printerName);

        /// <summary>取模板中已命名的数据字段列表。</summary>
        IReadOnlyList<string> GetSubStringNames();

        /// <summary>打印一张标签；失败时 <paramref name="error"/> 返回可读原因。</summary>
        bool PrintTag(string subStringName, string value, int copies, out string error);

        /// <summary>顺序批量打印。</summary>
        /// <param name="values">待打印的值序列（惰性枚举即可）。</param>
        /// <param name="subStringName">模板中的数据字段名。</param>
        /// <param name="copies">每行打印份数。</param>
        /// <param name="progressEvery">每多少行回报一次进度。</param>
        /// <param name="onProgress">进度回调（可为空）。</param>
        /// <param name="token">取消令牌。</param>
        PrintBatchResult PrintBatch(
            IEnumerable<string> values,
            string subStringName,
            int copies,
            int progressEvery,
            IProgress<PrintProgress> onProgress,
            CancellationToken token);
    }

    /// <summary>打印机查询（不依赖 Bartender SDK）。</summary>
    public static class PrinterHelper
    {
        /// <summary>系统中已安装的打印机名称。</summary>
        public static IReadOnlyList<string> GetInstalledPrinters()
        {
            var list = new List<string>();
            try
            {
                foreach (string name in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    if (!string.IsNullOrWhiteSpace(name)) list.Add(name);
                }
            }
            catch
            {
                // 读取打印子系统失败时返回空列表，由调用方给出提示
            }
            return list;
        }

        /// <summary>系统默认打印机名称（没有时返回 null）。</summary>
        public static string GetDefaultPrinter()
        {
            try
            {
                var settings = new System.Drawing.Printing.PrinterSettings();
                string name = settings.PrinterName;
                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
            catch
            {
                return null;
            }
        }
    }

#if BARTTENDER

    /// <summary>
    /// 基于 Seagull.BarTender.Print 的打印服务实现。
    ///
    /// 相比旧版 <c>PrintTag(string)</c> 的关键优化：
    ///   1. 引擎 + 模板常驻复用。打开一次 .btw 动辄几百毫秒，逐张标签「打开→打印→关闭」
    ///      是批量打印最大的瓶颈（旧代码每张都重开一遍）。
    ///   2. 守卫式提前返回，替代 4 层嵌套 if。
    ///   3. try/catch/finally 保证模板与引擎一定释放（旧代码抛异常时会漏掉 Close/Stop）。
    ///   4. <c>Cache.FlushInterval</c> 改到打印之前设置（旧代码写在打印之后，实际不生效）。
    ///   5. 打印前校验模板字段名是否存在，失败时给出可读原因，而不是抛 COMException。
    ///   6. 回传 Bartender 的 Messages，把「缺纸/离线/模板错误」等真实原因暴露出来。
    ///   7. 份数用 <c>IdenticalCopiesOfLabel</c>，由 SDK 一次出多份，比循环 N 次快得多。
    ///   8. 全部 SDK 调用加锁，兼容多线程调用。
    /// </summary>
    public sealed class BartenderPrintService : ILabelPrintService
    {
        private readonly object _sync = new object();
        private Engine _engine;
        private LabelFormatDocument _format;
        private string _templatePath;
        private bool _disposed;

        public BartenderPrintService()
        {
            PrintTimeoutMs = 30000;
        }

        /// <summary>单张标签的打印等待上限（毫秒），超时按失败处理。</summary>
        public int PrintTimeoutMs { get; set; }

        public bool IsTemplateOpen
        {
            get { lock (_sync) { return _format != null; } }
        }

        public string TemplatePath
        {
            get { lock (_sync) { return _templatePath; } }
        }

        public void OpenTemplate(string templatePath, string printerName)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("打印模板路径不能为空", nameof(templatePath));
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("打印模板不存在：" + templatePath, templatePath);

            lock (_sync)
            {
                ThrowIfDisposed();
                EnsureEngine();

                // 同一模板已打开：只切换打印机，不重复 Open（Open 是最慢的一步）
                if (_format != null &&
                    string.Equals(_templatePath, templatePath, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(printerName))
                        _format.PrintSetup.PrinterName = printerName;
                    return;
                }

                CloseTemplate();

                _format = _engine.Documents.Open(templatePath);
                _templatePath = templatePath;

                if (!string.IsNullOrWhiteSpace(printerName))
                    _format.PrintSetup.PrinterName = printerName;

                // 必须在打印之前设置：缓存刷新策略作用于后续作业
                _format.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
            }
        }

        public IReadOnlyList<string> GetSubStringNames()
        {
            lock (_sync)
            {
                if (_format == null) return new string[0];
                return _format.SubStrings.Select(s => s.Name).ToList();
            }
        }

        public bool PrintTag(string subStringName, string value, int copies, out string error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(subStringName))
            {
                error = "未指定模板数据字段名";
                return false;
            }

            lock (_sync)
            {
                if (_format == null)
                {
                    error = "打印模板尚未打开";
                    return false;
                }

                try
                {
                    SubString target = FindSubString(subStringName);
                    if (target == null)
                    {
                        error = "模板中不存在数据字段「" + subStringName + "」。模板现有字段："
                                + string.Join("、", GetSubStringNames());
                        return false;
                    }

                    target.Value = value ?? string.Empty;

                    if (_format.PrintSetup.SupportsIdenticalCopies)
                        _format.PrintSetup.IdenticalCopiesOfLabel = Math.Max(1, copies);

                    Messages messages;
                    Result result = _format.Print("标签打印", PrintTimeoutMs, out messages);
                    if (result == Result.Success)
                        return true;

                    error = DescribeResult(result, messages);
                    return false;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }
        }

        public PrintBatchResult PrintBatch(
            IEnumerable<string> values,
            string subStringName,
            int copies,
            int progressEvery,
            IProgress<PrintProgress> onProgress,
            CancellationToken token)
        {
            var result = new PrintBatchResult();
            if (values == null) return result;
            if (copies < 1) copies = 1;
            int every = Math.Max(1, progressEvery);

            int index = 0;
            foreach (string value in values)
            {
                if (token.IsCancellationRequested)
                {
                    result.Canceled = true;
                    break;
                }

                index++;

                string error;
                if (PrintTag(subStringName, value, copies, out error))
                {
                    result.PrintedRows++;
                    result.PrintedLabels += copies;
                }
                else
                {
                    result.FailedRows++;
                    result.Errors.Add("第 " + index + " 行（" + value + "）打印失败：" + error);
                }

                if (onProgress != null && (index == 1 || index % every == 0))
                {
                    onProgress.Report(new PrintProgress
                    {
                        Rows = result.PrintedRows + result.FailedRows,
                        Labels = result.PrintedLabels,
                        CurrentIndex = index
                    });
                }
            }

            if (onProgress != null)
            {
                onProgress.Report(new PrintProgress
                {
                    Rows = result.PrintedRows + result.FailedRows,
                    Labels = result.PrintedLabels,
                    CurrentIndex = index
                });
            }

            return result;
        }

        public void Dispose()
        {
            lock (_sync)
            {
                if (_disposed) return;
                _disposed = true;

                CloseTemplate();

                if (_engine != null)
                {
                    try { _engine.Stop(SaveOptions.DoNotSaveChanges); } catch { /* 引擎已失效时忽略 */ }
                    try { _engine.Dispose(); } catch { }
                    _engine = null;
                }
            }
        }

        // ------------------------------------------------------------------
        //  helpers
        // ------------------------------------------------------------------

        private void EnsureEngine()
        {
            if (_engine != null) return;

            try
            {
                // true = 构造后立即启动打印引擎
                _engine = new Engine(true);
                if (!_engine.IsAlive) _engine.Start();
            }
            catch (Exception ex)
            {
                _engine = null;
                throw new InvalidOperationException(
                    "启动 Bartender 打印引擎失败（请确认已安装 Bartender 且为可用授权版本）：" + ex.Message, ex);
            }
        }

        private void CloseTemplate()
        {
            if (_format == null) return;
            try
            {
                // 不保存对模板的修改
                _format.Close(SaveOptions.DoNotSaveChanges);
            }
            catch
            {
                // 模板已失效时忽略关闭异常
            }
            finally
            {
                _format = null;
                _templatePath = null;
            }
        }

        private SubString FindSubString(string name)
        {
            foreach (SubString item in _format.SubStrings)
            {
                if (string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
                    return item;
            }
            return null;
        }

        private static string DescribeResult(Result result, Messages messages)
        {
            var sb = new StringBuilder();

            if (result == Result.Timeout)
                sb.Append("打印超时：打印任务在限定时间内未完成。");
            else if (result == Result.Failure)
                sb.Append("打印失败。");
            else
                sb.Append("打印未成功（" + result + "）。");

            if (messages != null)
            {
                foreach (Seagull.BarTender.Print.Message message in messages)
                {
                    if (!string.IsNullOrEmpty(message.Text))
                        sb.AppendLine().Append(message.Text);
                }
            }

            return sb.ToString().Trim();
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(BartenderPrintService));
        }
    }

#else

    /// <summary>
    /// 未检测到 Bartender SDK 时的占位实现：编译仍可通过，运行时给出明确原因。
    /// </summary>
    public sealed class UnavailablePrintService : ILabelPrintService
    {
        internal const string Reason =
            "未检测到 Bartender SDK（Seagull.BarTender.Print.dll）。请安装 BarTender 后重新编译，" +
            "或在项目属性中手动添加 SDK 引用（程序集通常位于 BarTender 安装目录 SDK\\Assemblies 下）。";

        public bool IsTemplateOpen => false;

        public string TemplatePath => null;

        public void OpenTemplate(string templatePath, string printerName)
        {
            throw new NotSupportedException(Reason);
        }

        public IReadOnlyList<string> GetSubStringNames() => new string[0];

        public bool PrintTag(string subStringName, string value, int copies, out string error)
        {
            error = Reason;
            return false;
        }

        public PrintBatchResult PrintBatch(
            IEnumerable<string> values,
            string subStringName,
            int copies,
            int progressEvery,
            IProgress<PrintProgress> onProgress,
            CancellationToken token)
        {
            var result = new PrintBatchResult();
            result.FailedRows = 1;
            result.Errors.Add(Reason);
            return result;
        }

        public void Dispose()
        {
        }
    }

#endif

    /// <summary>打印服务工厂：有 SDK 用 SDK，没有 SDK 用占位实现。</summary>
    public static class PrintServiceFactory
    {
        public static ILabelPrintService Create()
        {
#if BARTTENDER
            return new BartenderPrintService();
#else
            return new UnavailablePrintService();
#endif
        }
    }
}

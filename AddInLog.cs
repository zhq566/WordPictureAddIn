using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace WordPictureAddIn
{
    /// <summary>
    /// 极简文件日志，用于排查加载项在某台机器上启动失败的原因。
    /// 日志写入 %LOCALAPPDATA%\WordPictureAddIn\addin.log。
    /// 只保留最近一次启动的内容：每次 Word 启动后第一次写日志时覆盖旧文件，
    /// 正常使用情况下内容基本一致，不需要累积历史。
    /// 所有写入都吞掉异常——日志本身绝不能影响加载项运行。
    /// </summary>
    internal static class AddInLog
    {
        /// <summary>单次会话内的日志上限，防止同一进程反复报错把文件撑爆。</summary>
        private const long MaxLogSizeBytes = 1024 * 1024;

        private static readonly object SyncRoot = new object();
        private static string _logFilePath;
        private static bool _disabled;

        /// <summary>本次进程是否已经写入过日志；首次写入需要覆盖上一次启动留下的内容。</summary>
        private static bool _sessionStarted;

        /// <summary>日志文件的完整路径。</summary>
        public static string LogFilePath
        {
            get
            {
                lock (SyncRoot)
                {
                    if (_logFilePath == null)
                    {
                        _logFilePath = Path.Combine(GetLogDirectory(), "addin.log");
                    }
                    return _logFilePath;
                }
            }
        }

        /// <summary>日志所在目录。</summary>
        public static string LogDirectory
        {
            get { return GetLogDirectory(); }
        }

        private static string GetLogDirectory()
        {
            try
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordPictureAddIn");
            }
            catch
            {
                return Path.Combine(Path.GetTempPath(), "WordPictureAddIn");
            }
        }

        /// <summary>写入一条信息行。</summary>
        public static void Write(string message)
        {
            WriteLine("INFO", message ?? string.Empty, null);
        }

        /// <summary>按格式写入一条信息行。</summary>
        public static void Write(string format, params object[] args)
        {
            string text;
            try
            {
                text = args == null || args.Length == 0
                    ? format
                    : string.Format(CultureInfo.InvariantCulture, format, args);
            }
            catch
            {
                text = format;
            }
            WriteLine("INFO", text, null);
        }

        /// <summary>写入一条警告行。</summary>
        public static void Warn(string message)
        {
            WriteLine("WARN", message ?? string.Empty, null);
        }

        /// <summary>写入一条异常行，包含完整堆栈。</summary>
        public static void Error(string context, Exception exception)
        {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(context))
            {
                builder.Append(context);
            }

            if (exception != null)
            {
                if (builder.Length > 0)
                {
                    builder.Append(" -> ");
                }
                builder.Append(exception.GetType().FullName);
                builder.Append(": ");
                builder.Append(exception.Message);

                var inner = exception.InnerException;
                var depth = 0;
                while (inner != null && depth < 5)
                {
                    builder.AppendLine();
                    builder.Append("  [inner] ");
                    builder.Append(inner.GetType().FullName);
                    builder.Append(": ");
                    builder.Append(inner.Message);
                    inner = inner.InnerException;
                    depth++;
                }

                if (!string.IsNullOrEmpty(exception.StackTrace))
                {
                    builder.AppendLine();
                    builder.Append(exception.StackTrace);
                }
            }

            WriteLine("ERROR", builder.ToString(), null);
        }

        private static void WriteLine(string level, string message, Exception ignored)
        {
            if (_disabled)
            {
                return;
            }

            lock (SyncRoot)
            {
                try
                {
                    var directory = GetLogDirectory();
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    var path = Path.Combine(directory, "addin.log");
                    var line = string.Concat(
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
                        " [", level, "] ",
                        message,
                        Environment.NewLine);

                    if (!_sessionStarted)
                    {
                        // 本次启动的第一条日志：覆盖上一次运行留下的内容，只保留最近一次。
                        File.WriteAllText(path, line, Encoding.UTF8);
                        _sessionStarted = true;
                        return;
                    }

                    // 同一进程内持续报错时兜底，避免文件无限增长。
                    var info = new FileInfo(path);
                    if (info.Exists && info.Length > MaxLogSizeBytes)
                    {
                        File.WriteAllText(path, line, Encoding.UTF8);
                        return;
                    }

                    File.AppendAllText(path, line, Encoding.UTF8);
                }
                catch
                {
                    // 任何 IO 失败都静默放弃，并且本次会话不再重复尝试
                    _disabled = true;
                }
            }
        }
    }
}
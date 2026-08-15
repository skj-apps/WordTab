using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace WordTab
{
    /// <summary>
    /// A deliberately dumb file log at %LOCALAPPDATA%\WordTab\wordtab.log.
    ///
    /// This is the primary way we see what happens inside WINWORD, where there is no console and
    /// a debugger is not always an option. Every method is best-effort and swallows its own
    /// errors: a logging failure must never be the thing that takes the add-in down, because Word
    /// responds to an add-in throwing during load by disabling it silently and permanently.
    /// </summary>
    internal static class Log
    {
        private const long MaxBytes = 512 * 1024;

        private static readonly object Gate = new object();
        private static readonly string Dir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordTab");
        private static readonly string Path_ = Path.Combine(Dir, "wordtab.log");

        internal static string FilePath { get { return Path_; } }

        internal static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(Dir);

                    // Roll rather than grow without bound. One file, no history: this log is for
                    // "what did the last run do", not an audit trail.
                    var info = new FileInfo(Path_);
                    if (info.Exists && info.Length > MaxBytes)
                        info.Delete();

                    string line = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0:yyyy-MM-dd HH:mm:ss.fff}  pid={1,-6} tid={2,-4}  {3}{4}",
                        DateTime.Now,
                        Process.GetCurrentProcess().Id,
                        Environment.CurrentManagedThreadId,
                        message,
                        Environment.NewLine);

                    File.AppendAllText(Path_, line, Encoding.UTF8);
                }
            }
            catch
            {
                // Intentionally empty. See the class comment.
            }
        }

        internal static void Error(string where, Exception ex)
        {
            Write("ERROR in " + where + ": " + (ex == null ? "(null)" : ex.ToString()));
        }
    }
}

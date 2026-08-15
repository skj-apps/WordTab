using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace WordTab
{
    /// <summary>
    /// The COM class Word instantiates. This is the entry point for everything WordTab ever does
    /// inside WINWORD.
    ///
    /// At this stage it only proves it loaded: it writes to the log and shows a one-shot banner.
    /// The window work (subclassing OpusApp, shrinking _WwF, painting the strip) comes next and
    /// hangs off OnStartupComplete / OnDisconnection.
    ///
    /// Rule for every callback below: nothing escapes. An exception thrown back across the COM
    /// boundary during load makes Word add us to its Resiliency\DisabledItems list, which is
    /// silent, sticky, and confusing to diagnose later.
    /// </summary>
    [ComVisible(true)]
    [Guid(Connect.ClsidString)]
    [ProgId(Connect.ProgIdString)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class Connect : IDTExtensibility2
    {
        // Single source of truth for the identity Word looks us up by.
        // install\install.ps1 parses these two literals out of this file and refuses to run if
        // they disagree with its own copy, so the code and the registration cannot drift apart.
        public const string ClsidString = "4BF75ED9-10EE-4866-BF4A-3D663A4149A1";
        public const string ProgIdString = "WordTab.Connect";

        private const string SettingsKey = @"Software\WordTab";

        private object _wordApp;
        private object _addInInst;

        /// <summary>Static, not instance: Word can create more than one instance of us over a
        /// process lifetime, and one banner per process is the point.</summary>
        private static int _bannerShown;

        public Connect()
        {
            // Deliberately the first thing logged. If activation itself is broken we will see this
            // line and nothing after it; if registration is broken we will see nothing at all.
            // That distinction is the whole diagnosis.
            Log.Write("---- Connect ctor ---- " + DescribeSelf());
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _wordApp = application;
                _addInInst = addInInst;
                Log.Write("OnConnection  mode=" + connectMode + "  " + DescribeHost(application));
                ShowBannerOnce(connectMode);
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        public void OnStartupComplete(ref Array custom)
        {
            // Word's UI exists by now. This is where the window work will start, which is why the
            // main window handle is logged here — it is the OpusApp the next slice has to subclass.
            try
            {
                Log.Write("OnStartupComplete  mainWindow=" + MainWindowHandle());
            }
            catch (Exception ex) { Log.Error("OnStartupComplete", ex); }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            try { Log.Write("OnAddInsUpdate"); }
            catch (Exception ex) { Log.Error("OnAddInsUpdate", ex); }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try { Log.Write("OnBeginShutdown"); }
            catch (Exception ex) { Log.Error("OnBeginShutdown", ex); }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                Log.Write("OnDisconnection  mode=" + removeMode);

                // ext_dm_UserClosed means Word keeps running without us, so teardown has to be
                // real. Once this class owns window state, undoing it belongs here and must not
                // depend on the process exiting.
                Release(ref _addInInst);
                Release(ref _wordApp);
            }
            catch (Exception ex) { Log.Error("OnDisconnection", ex); }
        }

        // ---- diagnostics -------------------------------------------------------------------

        /// <summary>Which copy of us actually loaded, and under which CLR. Proves the DLL came
        /// from the install directory and not some stale build left somewhere else.</summary>
        private static string DescribeSelf()
        {
            var sb = new StringBuilder();
            try
            {
                Assembly me = typeof(Connect).Assembly;
                sb.Append("asm=").Append(me.FullName);
                sb.Append("  from=").Append(me.Location);
                sb.Append("  clr=").Append(Environment.Version);
                sb.Append("  host=").Append(Process.GetCurrentProcess().MainModule.FileName);
                sb.Append("  bits=").Append(IntPtr.Size * 8);
            }
            catch (Exception ex) { sb.Append("  (describe failed: ").Append(ex.Message).Append(')'); }
            return sb.ToString();
        }

        /// <summary>Interrogate the Application object Word handed us, late-bound. This is the
        /// proof that we are talking to the real Word rather than merely being instantiated:
        /// the build number should match the one recorded for this rig.</summary>
        private static string DescribeHost(object application)
        {
            if (application == null) return "app=(null)";

            var sb = new StringBuilder("app=");
            sb.Append(GetProperty(application, "Name") ?? "(?)");
            sb.Append(" version=").Append(GetProperty(application, "Version") ?? "(?)");
            sb.Append(" build=").Append(GetProperty(application, "Build") ?? "(?)");

            object docs = GetPropertyObject(application, "Documents");
            if (docs != null)
            {
                sb.Append(" documents=").Append(GetProperty(docs, "Count") ?? "(?)");
                Marshal.ReleaseComObject(docs);
            }
            return sb.ToString();
        }

        private static string GetProperty(object target, string name)
        {
            object value = GetPropertyObject(target, name);
            if (value == null) return null;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static object GetPropertyObject(object target, string name)
        {
            try
            {
                return target.GetType().InvokeMember(
                    name, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture);
            }
            catch
            {
                // Word not answering a property is information, not a failure. Callers print "(?)".
                return null;
            }
        }

        private static string MainWindowHandle()
        {
            try { return "0x" + Process.GetCurrentProcess().MainWindowHandle.ToString("X"); }
            catch { return "(?)"; }
        }

        private static void Release(ref object comObject)
        {
            object o = comObject;
            comObject = null;
            if (o == null || !Marshal.IsComObject(o)) return;
            try { Marshal.ReleaseComObject(o); } catch { }
        }

        // ---- load banner -------------------------------------------------------------------

        /// <summary>
        /// The visible proof that we loaded. Shown on a background thread on purpose: a modal
        /// dialog on Word's UI thread during startup would block Word until it is dismissed, and
        /// an add-in that can hang the host while proving itself is not proving much.
        /// Switch it off once it has served its purpose with install.ps1 -NoBanner.
        /// </summary>
        private static void ShowBannerOnce(ext_ConnectMode connectMode)
        {
            if (!BannerEnabled()) return;
            if (Interlocked.Exchange(ref _bannerShown, 1) != 0) return;

            string text =
                "WordTab is loaded inside Word." + Environment.NewLine + Environment.NewLine +
                "connect mode: " + connectMode + Environment.NewLine +
                "log: " + Log.FilePath;

            var thread = new Thread(delegate ()
            {
                try { MessageBox(IntPtr.Zero, text, "WordTab", MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND | MB_TOPMOST); }
                catch (Exception ex) { Log.Error("banner", ex); }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private static bool BannerEnabled()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKey))
                {
                    if (key == null) return true;
                    object value = key.GetValue("ShowLoadBanner", 1);
                    return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
                }
            }
            catch { return true; }
        }

        private const uint MB_OK = 0x00000000;
        private const uint MB_ICONINFORMATION = 0x00000040;
        private const uint MB_SETFOREGROUND = 0x00010000;
        private const uint MB_TOPMOST = 0x00040000;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.IO;

namespace IEManipulator
{
    public partial class MainForm : Form
    {
        private GlobalKeyboardHook _globalKeyboardHook;

        public void SetupKeyboardHooks()
        {
            _globalKeyboardHook = new GlobalKeyboardHook();
            _globalKeyboardHook.KeyboardPressed += OnKeyPressed;
        }

        private void OnKeyPressed(object sender, GlobalKeyboardHookEventArgs e)
        {
            if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.KeyDown && e.KeyboardData.VirtualCode == 80)
            {
                //Start();

                e.Handled = true;
            }
        }

        public MainForm()
        {
            InitializeComponent();

            SetupKeyboardHooks();
        }

        static private SHDocVw.ShellWindowsClass shellWindows = new
        SHDocVw.ShellWindowsClass();

        private void Form1_Load(object sender, EventArgs e)
        {
            /*oreach (SHDocVw.InternetExplorer ie in shellWindows)
            {
                //MessageBox.Show("ie.Location:" + ie.LocationURL);
                //ie.BeforeNavigate2 += new
                //SHDocVw.DWebBrowserEvents2_BeforeNavigate2EventHandler(this.ie_BeforeNavigate2);

                try
                {
                    var dom = ie.Document as mshtml.HTMLDocumentClass;

                    if (dom != null)
                    {
                        var text = dom.documentElement.outerHTML;

                        if (!string.IsNullOrEmpty(text))
                        {
                            string filePath = Path.Combine(Application.StartupPath, "Source.txt");

                            using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                            {
                                writer.Write(text);
                            }

                            MessageBox.Show("Internet Explorer içeriği Source.txt sayfasına yazıldı.");

                            Application.Exit();
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
                }
            }*/
        }

        public void ie_BeforeNavigate2(object pDisp, ref object url, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            MessageBox.Show("event received!");
        }

        /*protected override CreateParams CreateParams
        {
            get
            {
                var Params = base.CreateParams;
                Params.ExStyle |= 0x80;
                return Params;
            }
        }*/

        private void button1_Click(object sender, EventArgs e)
        {
            Start();
        }

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public void Start()
        {
            Process[] processList = Process.GetProcesses();

            foreach (Process P in processList)
            {
                if (P.ProcessName.Equals("iexplore"))
                {

                    IntPtr edit = P.MainWindowHandle;

                    ForceWindowToForeground(edit);

                    if (GetForegroundWindow() == edit)
                    {

                        SendKeys.SendWait("^u");
                        SendKeys.Flush();
                    }
                }
            }
        }

        public static void AttachedThreadInputAction(Action action)
        {
            var foreThread = GetWindowThreadProcessId(GetForegroundWindow(),
                IntPtr.Zero);
            var appThread = GetCurrentThreadId();
            bool threadsAttached = false;
            try
            {
                threadsAttached =
                    foreThread == appThread ||
                    AttachThreadInput(foreThread, appThread, true);
                if (threadsAttached) action();
                else throw new ThreadStateException("AttachThreadInput failed.");
            }
            finally
            {
                if (threadsAttached)
                    AttachThreadInput(foreThread, appThread, false);
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // When you don't want the ProcessId, use this overload and pass 
        // IntPtr.Zero for the second parameter
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd,
            IntPtr ProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        /// The GetForegroundWindow function returns a handle to the 
        /// foreground window.
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach,
            uint idAttachTo, bool fAttach);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(HandleRef hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        public const uint SW_SHOW = 5;
        ///<summary>
        /// Forces the window to foreground.
        ///</summary>
        ///hwnd”>The HWND.</param>
        public static void ForceWindowToForeground(IntPtr hwnd)
        {
            AttachedThreadInputAction(
                () =>
                {
                    BringWindowToTop(hwnd);
                    ShowWindow(hwnd, SW_SHOW);
                });
        }
        public static IntPtr SetFocusAttached(IntPtr hWnd)
        {
            var result = new IntPtr();
            AttachedThreadInputAction(
                () =>
                {
                    result = SetFocus(hWnd);
                });
            return result;
        }

        private void MainForm_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _globalKeyboardHook?.Dispose();
        }
    }
}

using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WinCC_Assistant
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private Timer timer;

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_COMMAND = 0x0111;
        private const int OK_BUTTON_ID = 1; // 确认按钮的 ID
        private const string WinCCInfoWindow = "WinCC Information";
        private const string WinCCInfoWindowCN = "WinCC 信息";
        private const string PdlRtProcessName = "PdlRt";
        private const string GscrtProcessName = "gscrt";
        private const string GscrtPath = @"C:\Program Files (x86)\Siemens\Automation\SCADA-RT_V11\WinCC\bin\gscrt.exe";

        public Form1()
        {
            InitializeComponent();
            this.Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("WinCC_Assistant.icon.ico"));
            InitializeTrayIcon();
            InitializeTimer();
            HideForm();
        }

        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Designed by Amaury", OnShowInfo);
            trayMenu.MenuItems.Add("-");
            trayMenu.MenuItems.Add("Exit", OnExit);

            trayIcon = new NotifyIcon
            {
                Text = "WinCC Assistant",
                Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("WinCC_Assistant.icon.ico")),
                ContextMenu = trayMenu,
                Visible = true
            };

            trayIcon.MouseClick += TrayIcon_MouseClick;
        }

        private void InitializeTimer()
        {
            timer = new Timer { Interval = 1000 }; // 每1秒执行一次
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void HideForm()
        {
            this.WindowState = FormWindowState.Minimized; // 隐藏窗体
            this.ShowInTaskbar = false; // 不显示在任务栏
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {

        }


        private void Timer_Tick(object sender, EventArgs e)
        {
            // 查找窗口并发送消息
            FindAndSendMessage(WinCCInfoWindow);
            FindAndSendMessage(WinCCInfoWindowCN);

            if (CheckProcessExists(PdlRtProcessName) && !CheckProcessExists(GscrtProcessName))
            {
                StartProcess(GscrtPath);
            }
        }

        private void FindAndSendMessage(string windowName)
        {
            IntPtr hwnd = FindWindow(null, windowName);
            if (hwnd != IntPtr.Zero)
            {
                // 发送确认按钮消息
                SendMessage(hwnd, WM_COMMAND, (IntPtr)OK_BUTTON_ID, IntPtr.Zero);
            }
        }

        private bool CheckProcessExists(string processName)
        {
            return Process.GetProcesses().Any(p => p.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase));
        }

        private void StartProcess(string path)
        {
            try
            {
                using (var process = new Process())
                {
                    process.StartInfo.FileName = path;
                    process.Start();
                }
            }
            catch (Exception)
            {

            }
        }
        private void OnShowInfo(object sender, EventArgs e)
        {
            MessageBox.Show("1. Automatically close the WinCC Information pop-up window\n" +
                            "2. Processes guarding WinCC scheduled tasks",
                            "WinCC Assistant",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }
        private void OnExit(object sender, EventArgs e)
        {
            timer.Stop(); // 停止计时器
            trayIcon.Visible = false; // 隐藏托盘图标
            Application.Exit();       // 退出应用程序
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Hide(); // 加载时隐藏窗体
        }
    }
}
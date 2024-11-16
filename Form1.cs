using CCWSTLib;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinCC_Assistant
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;  // 系统托盘图标
        private ContextMenu trayMenu;  // 托盘图标右键菜单
        private IntPtr winEventHook;  // 用于监听窗口事件的钩子
        private ManagementEventWatcher PdlRtStopWatcher;  // 用于监控进程停止事件
        private ManagementEventWatcher GscrtStopWatcher;  // 用于监控进程停止事件

        private static CWinCCStart WinCC; // WinCC 对象

        // WinEventDelegate 委托，定义了当窗口事件发生时的回调函数
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        private WinEventDelegate eventDelegate;

        // 常量定义
        private const uint EVENT_OBJECT_CREATE = 0x8000;  // 窗口创建事件
        private const uint WINEVENT_OUTOFCONTEXT = 0;  // 事件标志，表示不在上下文中
        private const uint WM_COMMAND = 0x0111;  // 用于发送窗口命令的消息
        private const int OK_BUTTON_ID = 1;  // 确认按钮的 ID
        private static readonly string[] WinCCInfoWindows = { "WinCC Information", "WinCC 信息" };  // 需要关闭的 WinCC 信息窗口标题

        // 进程相关常量
        private const string PdlRtProcessName = "PdlRt";  // 需要监控的 PdlRt 进程名称
        private const string GscrtProcessName = "gscrt";  // 需要监控的 Gscrt 进程名称
        private const string PdlRtPath = @"C:\Program Files (x86)\Siemens\Automation\SCADA-RT_V11\WinCC\bin\PdlRt.exe";  // PdlRt 进程的路径
        private const string GscrtPath = @"C:\Program Files (x86)\Siemens\Automation\SCADA-RT_V11\WinCC\bin\gscrt.exe";  // Gscrt 进程的路径
        private const string ResetWinCCPath = @"C:\Program Files (x86)\Siemens\Automation\SCADA-RT_V11\WinCC\bin\Reset_WinCC.vbs";  // Reset_WinCC 进程的路径

        private const int RetryDelayMilliseconds = 5000;  // 重试间隔（毫秒）

        // 图标
        private static readonly Icon appIcon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("WinCC_Assistant.icon.ico"));

        // Windows API 调用，用于设置和取消 WinEvent 钩子
        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        // 用于发送消息到窗口
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // 获取窗口标题
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        // 查找窗口
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public Form1()
        {
            InitializeComponent();
            this.Icon = appIcon;  // 设置窗体图标
            InitializeTrayIcon();  // 初始化系统托盘图标
            InitializeWinEventHook();  // 初始化窗口事件钩子
            InitializeProcessWatcher();  // 初始化进程监控
            CloseExistingPopups();  // 关闭已有的弹窗
            HideForm();  // 隐藏窗体
        }

        // 初始化托盘图标和右键菜单
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Designed by Amaury");
            trayMenu.MenuItems.Add("-");  // 分隔线
            trayMenu.MenuItems.Add("Reset WinCC", async (sender, e) => await ResetWinCC());  // 菜单项：重置WinCC
            trayMenu.MenuItems.Add("-");  // 分隔线
            trayMenu.MenuItems.Add("Exit", OnExit);  // 菜单项：退出

            trayIcon = new NotifyIcon
            {
                Text = "WinCC Assistant",  // 鼠标悬停时显示的文本
                Icon = appIcon,  // 设置图标
                ContextMenu = trayMenu,  // 设置右键菜单
                Visible = true  // 显示托盘图标
            };

            trayIcon.MouseClick += TrayIcon_MouseClick;  // 设置托盘图标点击事件
        }

        // 隐藏窗体，使其最小化到系统托盘
        private void HideForm()
        {
            this.WindowState = FormWindowState.Minimized;  // 设置窗体最小化
            this.ShowInTaskbar = false;  // 不显示在任务栏
        }

        // 托盘图标点击事件
        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ShowBalloonNotification("1. Automatically close the WinCC Information pop-up window\n" +
                            "2. Monitor and guard WinCC processes");  // 显示气泡通知
            }
        }

        // 显示系统托盘气泡通知
        private void ShowBalloonNotification(string message)
        {
            trayIcon.ShowBalloonTip(1000, "", message, ToolTipIcon.None);  // 显示气泡通知，持续1秒
        }

        // 设置 WinEventHook 监听窗口创建事件
        private void InitializeWinEventHook()
        {
            eventDelegate = new WinEventDelegate(WinEventProc);  // 设置回调函数
            winEventHook = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_CREATE, IntPtr.Zero, eventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);  // 注册钩子
        }

        // 回调函数，当指定窗口创建时被调用
        private void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            string windowTitle = GetWindowTitle(hwnd);  // 获取窗口标题
            if (Array.Exists(WinCCInfoWindows, title => title.Equals(windowTitle, StringComparison.OrdinalIgnoreCase)))
            {
                SendMessage(hwnd, WM_COMMAND, (IntPtr)OK_BUTTON_ID, IntPtr.Zero);  // 点击确认按钮
            }
        }

        // 获取窗口标题
        private string GetWindowTitle(IntPtr hwnd)
        {
            StringBuilder title = new StringBuilder(256);
            GetWindowText(hwnd, title, title.Capacity);  // 获取窗口标题
            return title.ToString();
        }

        // 检测已有弹窗并关闭
        private void CloseExistingPopups()
        {
            foreach (var windowTitle in WinCCInfoWindows)
            {
                ClosePopupWindow(windowTitle);  // 关闭每一个符合条件的弹窗
            }
        }

        private void ClosePopupWindow(string windowTitle)
        {
            IntPtr hwnd = FindWindow(null, windowTitle);  // 查找窗口
            if (hwnd != IntPtr.Zero)
            {
                SendMessage(hwnd, WM_COMMAND, (IntPtr)OK_BUTTON_ID, IntPtr.Zero);  // 发送点击确认按钮的消息
            }
        }

        // 初始化进程监控，监控进程的停止
        private void InitializeProcessWatcher()
        {
            // 初始化 WinCC 对象
            WinCC = new CWinCCStart();

            WqlEventQuery PdlRtStopQuery = new WqlEventQuery($"SELECT * FROM __InstanceDeletionEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = '{PdlRtProcessName}.exe'");
            WqlEventQuery GscrtStopQuery = new WqlEventQuery($"SELECT * FROM __InstanceDeletionEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = '{GscrtProcessName}.exe'");

            PdlRtStopWatcher = new ManagementEventWatcher(PdlRtStopQuery);
            PdlRtStopWatcher.EventArrived += async (sender, e) => await OnProcessStop(PdlRtPath);
            PdlRtStopWatcher.Start();  // 启动监视器

            GscrtStopWatcher = new ManagementEventWatcher(GscrtStopQuery);
            GscrtStopWatcher.EventArrived += async (sender, e) => await OnProcessStop(GscrtPath);
            GscrtStopWatcher.Start();  // 启动监视器
        }

        // 当检测到目标进程终止时，检查项目状态并重启进程
        private async Task OnProcessStop(string processPath)
        {
            // 获取当前项目状态
            WinCC.GetProjectStatus(out _project_status status, out string _);

            // 检查项目状态是否为 PROJECT_STATUS_ACTIVATE_END
            if (status == _project_status.PROJECT_STATUS_ACTIVATE_END)
            {
                await StartProcess(processPath); // 启动目标进程
            }
        }

        // 启动指定路径的进程，并无限重试
        private async Task StartProcess(string path)
        {
            try
            {
                Process.Start(path);  // 启动进程
            }
            catch (Exception)
            {
                // 启动失败，继续重试
                await Task.Delay(RetryDelayMilliseconds);  // 等待一段时间再重试
            }
        }

        // 调用WinCC Reset脚本
        private async Task ResetWinCC()
        {
            await StartProcess(ResetWinCCPath); // 启动目标进程
        }
        // 退出应用程序
        private void OnExit(object sender, EventArgs e)
        {
            PdlRtStopWatcher?.Stop();  // 停止监视器
            GscrtStopWatcher?.Stop();  // 停止监视器
            UnhookWinEvent(winEventHook);  // 注销钩子
            trayIcon.Visible = false;  // 隐藏托盘图标
            Application.Exit();  // 退出应用程序
        }

        // 窗体加载时自动隐藏
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Hide();  // 隐藏窗体
        }
    }
}

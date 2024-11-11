using System;
using System.Windows.Forms;

namespace WinCC_Assistant
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            using (var mutex = new System.Threading.Mutex(true, "WinCC_Assistant", out bool isNewInstance))
            {
                if (!isNewInstance)
                {
                    return; // 如果已经有实例在运行，直接退出
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
        }
    }
}
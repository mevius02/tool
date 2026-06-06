using ColorPicker;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal static class Program
{
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    [STAThread]
    static void Main()
    {
        SetProcessDPIAware(); // ズレ解消

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new Form1());
    }
}

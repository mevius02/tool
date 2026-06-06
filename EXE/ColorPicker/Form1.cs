using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ColorPicker
{
    public partial class Form1 : Form
    {
        private string currentHex = "#FFFFFF";
        private Panel panelColor;

        // --- Windows API ---
        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("gdi32.dll")]
        static extern uint GetPixel(IntPtr hdc, int x, int y);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const uint SWP_NOSIZE = 0x0001;
        const uint SWP_NOMOVE = 0x0002;
        const uint SWP_SHOWWINDOW = 0x0040;

        public struct POINT { public int X; public int Y; }

        // --- グローバルマウスフック ---
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;

        private static IntPtr hookID = IntPtr.Zero;
        private LowLevelMouseProc hookCallback;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetModuleHandle(string lpModuleName);

        // --- Timer & Magnifier ---
        Timer timer = new Timer();
        Form magnifierForm = new Form();
        PictureBox magnifierBox = new PictureBox();

        // コンストラクタ
        public Form1()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.ShowIcon = false;
            this.ControlBox = false;
            this.BackColor = Color.White;

            // Formサイズ
            this.Size = new Size(120, 50);

            // 色パネル
            panelColor = new Panel();
            panelColor.BackColor = Color.Black;
            panelColor.Paint += PanelColor_Paint;
            this.Controls.Add(panelColor);

            // レイアウト
            this.Resize += (s, e) => LayoutColorPanel();
            LayoutColorPanel();

            // 右上固定
            this.Load += (s, e) =>
            {
                this.Left = Screen.PrimaryScreen.WorkingArea.Width - this.Width;
                this.Top = 0;
            };

            // フック開始
            hookCallback = HookCallback;
            hookID = SetHook(hookCallback);

            // 虫眼鏡
            magnifierForm.FormBorderStyle = FormBorderStyle.None;
            magnifierForm.TopMost = true;
            magnifierForm.ShowInTaskbar = false;
            magnifierForm.Width = 150;
            magnifierForm.Height = 150;
            magnifierForm.BackColor = Color.Black;

            magnifierBox.Dock = DockStyle.Fill;
            magnifierBox.SizeMode = PictureBoxSizeMode.StretchImage;
            magnifierForm.Controls.Add(magnifierBox);

            magnifierForm.Show();

            timer.Interval = 30;
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void LayoutColorPanel()
        {
            int margin = 4;

            int w = (int)(this.Width * 0.95);
            int h = (int)(this.Height * 0.90);

            int x = (this.Width - w) / 2;
            int y = (this.Height - h) / 2;

            panelColor.Location = new Point(x, y);
            panelColor.Size = new Size(w, h);

            panelColor.Padding = new Padding(margin);
            panelColor.BackColor = Color.White;
        }

        private void PanelColor_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            string text = currentHex;
            Font font = new Font("Consolas", 14, FontStyle.Bold);

            SizeF size = g.MeasureString(text, font);
            float x = (panelColor.Width - size.Width) / 2;
            float y = (panelColor.Height - size.Height) / 2;

            using (Brush outline = new SolidBrush(Color.Black))
            {
                g.DrawString(text, font, outline, x - 1, y);
                g.DrawString(text, font, outline, x + 1, y);
                g.DrawString(text, font, outline, x, y - 1);
                g.DrawString(text, font, outline, x, y + 1);
            }

            using (Brush brush = new SolidBrush(Color.White))
            {
                g.DrawString(text, font, brush, x, y);
            }
        }

        private IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                Clipboard.SetText(currentHex);
                UnhookWindowsHookEx(hookID);
                magnifierForm.Hide();
                this.Close();
                return (IntPtr)1;
            }

            return CallNextHookEx(hookID, nCode, wParam, lParam);
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            POINT p;
            GetCursorPos(out p);

            Bitmap bmp = new Bitmap(20, 20);
            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                gfx.CopyFromScreen(p.X - 10, p.Y - 10, 0, 0, new Size(20, 20));
            }

            using (Graphics cross = Graphics.FromImage(bmp))
            using (Pen pen = new Pen(Color.Red, 1))
            {
                cross.DrawLine(pen, 10, 0, 10, 20);
                cross.DrawLine(pen, 0, 10, 20, 10);
            }

            magnifierBox.Image = bmp;

            int magW = magnifierForm.Width;
            int magH = magnifierForm.Height;
            int offset = 20;

            int mx = p.X + offset;
            int my = p.Y + offset;

            int sw = Screen.PrimaryScreen.WorkingArea.Width;
            int sh = Screen.PrimaryScreen.WorkingArea.Height;

            if (mx + magW > sw) mx = p.X - magW - offset;
            if (my + magH > sh) my = p.Y - magH - offset;
            if (mx < 0) mx = p.X + offset;
            if (my < 0) my = p.Y + offset;

            magnifierForm.Left = mx;
            magnifierForm.Top = my;

            SetWindowPos(magnifierForm.Handle, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            IntPtr hdc = GetDC(IntPtr.Zero);
            uint pixel = GetPixel(hdc, p.X, p.Y);

            int r = (int)(pixel & 0x000000FF);
            int g = (int)((pixel & 0x0000FF00) >> 8);
            int b = (int)((pixel & 0x00FF0000) >> 16);

            Color c = Color.FromArgb(r, g, b);

            currentHex = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
            panelColor.BackColor = c;
            panelColor.Invalidate();
        }
    }
}

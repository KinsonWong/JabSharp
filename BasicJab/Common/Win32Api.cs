using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using Win32;
using System.Threading.Tasks;

namespace BasicJab.Common
{
    public static class Win32Api
    {
        //移动鼠标 
        const int MOUSEEVENTF_MOVE = 0x0001;
        //模拟鼠标左键按下 
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        //模拟鼠标左键抬起 
        const int MOUSEEVENTF_LEFTUP = 0x0004;
        //模拟鼠标右键按下 
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        //模拟鼠标右键抬起 
        const int MOUSEEVENTF_RIGHTUP = 0x0010;
        //模拟鼠标中键按下 
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        //模拟鼠标中键抬起 
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        //标示是否采用绝对坐标 
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        [DllImport("user32.dll", EntryPoint = "ShowWindow")]
        private extern static IntPtr ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private delegate bool EnumWindowsProc(IntPtr hWnd, int lParam);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        private static extern IntPtr GetShellWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        private static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int Y, int width, int height, int wFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool LogicalToPhysicalPointForPerMonitorDPI(IntPtr hwnd, ref POINT lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

        [DllImport("gdi32.dll")]
        static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);

        [DllImport("user32.dll")]
        static extern int GetWindowRgn(IntPtr hWnd, IntPtr hRgn);

        [DllImport("gdi32.dll")]
        static extern bool DeleteObject(IntPtr hObject);

        [DllImport("user32")]
        static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("User32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("User32")]
        static extern bool OpenClipboard(IntPtr hwnd);

        [DllImport("User32")]
        static extern bool CloseClipboard();

        [DllImport("User32")]
        static extern bool EmptyClipboard();

        [DllImport("User32")]
        static extern bool IsClipboardFormatAvailable(int format);

        [DllImport("User32")]
        static extern IntPtr GetClipboardData(int uFormat);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern IntPtr SetClipboardData(int uFormat, IntPtr hMem);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

        [DllImport("Shcore.dll")]
        private static extern IntPtr GetDpiForMonitor([In]IntPtr hmonitor, [In]DpiType dpiType, [Out]out uint dpiX, [Out]out uint dpiY);

        private enum DpiType
        {
            Effective = 0,
            Angular = 1,
            Raw = 2,
        }
        
        /// <summary>
        /// 另一种方式休眠
        /// </summary>
        /// <param name="milliseconds"></param>
        private static void MySleep(int milliseconds)
        {
            Task.Delay(milliseconds).Wait();
        }

        /// <summary>
        /// 枚举所有已打开的windows窗口，窗口标题和句柄 写入字典
        /// </summary>
        /// <returns></returns>
        public static IDictionary<IntPtr, string> GetAllOpenWindows()
        {
            IntPtr shellWindow = GetShellWindow();
            Dictionary<IntPtr, string> windows = new Dictionary<IntPtr, string>();

            EnumWindows(delegate (IntPtr hWnd, int lParam)
            {
                if (hWnd == shellWindow) return true;
                if (!IsWindowVisible(hWnd)) return true;

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return true;

                StringBuilder builder = new StringBuilder(length);
                GetWindowText(hWnd, builder, length + 1);

                windows[hWnd] = builder.ToString();
                return true;

            }, 0);

            return windows;
        }


        /// <summary>
        /// 控制指定句柄的窗口显示大小
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="nCmdShow"></param>
        /// <returns></returns>
        public static IntPtr _ShowWindow(IntPtr hwnd, int nCmdShow)
        {
            return ShowWindow(hwnd, nCmdShow);
        }

        /// <summary>
        /// 用标题找窗口句柄
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public static IntPtr FindWindowByCaption(string caption)
        {
            return FindWindow(null, caption);
        }

        /// <summary>
        /// 将指定句柄的窗口放到顶层
        /// </summary>
        /// <param name="hwnd"></param>
        public static void SetTopWindow(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
        }

        /// <summary>
        /// 用句柄获取进程PID
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static int GetPidFromHwnd(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return 0;
            int pid;
            GetWindowThreadProcessId(hwnd, out pid);
            return pid;
        }

        /// <summary>
        /// 设置窗口大小
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void SetWindowSize(IntPtr hwnd, int width, int height)
        {
            const short SWP_NOMOVE = 0X2;
            const short SWP_NOZORDER = 0X4;
            SetWindowPos(hwnd, 0, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
        }

        /// <summary>
        /// 设置窗口位置
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public static void SetWindowPosition(IntPtr hwnd, int x, int y)
        {
            const short SWP_NOZORDER = 0X4;
            const short SWP_NOSIZE = 0X1;
            SetWindowPos(hwnd, 0, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
        }

        /// <summary>
        /// 获取指定句柄的窗口的截图，返回bitmap对象
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static Bitmap GetScreenshot(IntPtr hwnd)
        {
            SetTopWindow(hwnd);

            RECT rc;
            GetWindowRect(new HandleRef(null, hwnd), out rc);

            Bitmap bmp = new Bitmap(rc.Right - rc.Left, rc.Bottom - rc.Top, PixelFormat.Format32bppArgb);
            Graphics gfxBmp = Graphics.FromImage(bmp);
            IntPtr hdcBitmap;
            try
            {
                hdcBitmap = gfxBmp.GetHdc();
            }
            catch
            {
                return null;
            }
            bool succeeded = PrintWindow(hwnd, hdcBitmap, 0);
            gfxBmp.ReleaseHdc(hdcBitmap);
            if (!succeeded)
            {
                gfxBmp.FillRectangle(new SolidBrush(Color.Gray), new Rectangle(Point.Empty, bmp.Size));
            }
            IntPtr hRgn = CreateRectRgn(0, 0, 0, 0);
            GetWindowRgn(hwnd, hRgn);
            Region region = Region.FromHrgn(hRgn);//err here once
            DeleteObject(hRgn);
            if (!region.IsEmpty(gfxBmp))
            {
                gfxBmp.ExcludeClip(region);
                gfxBmp.Clear(Color.Transparent);
            }
            gfxBmp.Dispose();
            return bmp;
        }

        public static RECT GetRectFromHwnd(IntPtr hwnd)
        {
            RECT rc;
            GetWindowRect(new HandleRef(null, hwnd), out rc);
            return rc;
        }

        /// <summary>
        /// 调用win api 模拟鼠标点击
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="holdSeconds"></param>
        /// <param name="btn"></param>
        public static void Mouse_Click(int x, int y, int holdSeconds = 0, string btn = "left")
        {
            int mouseDownAct, mouseUpAct;
            if (btn == "left")
            {
                mouseDownAct = MOUSEEVENTF_LEFTDOWN;
                mouseUpAct = MOUSEEVENTF_LEFTUP;
            }
            else
            {
                mouseDownAct = MOUSEEVENTF_RIGHTDOWN;
                mouseUpAct = MOUSEEVENTF_RIGHTUP;
            }

            SetCursorPos(x, y);
            mouse_event(mouseDownAct, x, y, 0, 0);
            if (holdSeconds > 0) MySleep(holdSeconds * 1000);
            mouse_event(mouseUpAct, x, y, 0, 0);
        }

        /// <summary>
        /// 移动鼠标到指定位置
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public static void Move_Cursor(int x, int y)
        {
            SetCursorPos(x, y);
        }


        /// <summary>
        /// 设置剪切板文字
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool SetClipboardText(string text)
        {
            if (OpenClipboard(IntPtr.Zero))
            {
                IntPtr tmp = Marshal.StringToHGlobalUni(text);
                SetClipboardData(13, tmp);
                CloseClipboard();
                Marshal.FreeHGlobal(tmp);
                return true;
            }
            else { return false; }
        }

        /// <summary>
        /// 清空剪切板
        /// </summary>
        /// <returns></returns>
        public static bool ClearClipboard()
        {
            if (OpenClipboard(IntPtr.Zero))
            {
                EmptyClipboard();
                CloseClipboard();
                return true;
            }
            else { return false; }
        }


        /// <summary>
        /// 逻辑坐标转化为物理坐标
        /// </summary>
        /// <param name="xValue"></param>
        /// <param name="yValue"></param>
        public static void ConvertPoint_LogicalToPhysical(ref int xValue, ref int yValue)
        {
            POINT p = new POINT
            {
                x = xValue,
                y = yValue
            };

            //计算缩放比列  
            IntPtr pmScrenn=MonitorFromPoint(p, 1);
            uint dpiX,dpiY;

            GetDpiForMonitor(pmScrenn, DpiType.Effective, out dpiX, out dpiY);

            double scalingFactor = dpiY / 96.0;

            //逻辑坐标乘以缩放比列得到物理坐标
            xValue = (int)(xValue*scalingFactor);
            yValue = (int)(yValue * scalingFactor);
        }
    }
}

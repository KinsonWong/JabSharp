using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BasicJab.ComInterface;
using BasicJab.Common;

namespace BasicJab
{
    //空间+类名 记录到注册表中，给其他语言创建对象用
    [ProgId("BasicJab.IUtility")]
    [Guid("4786A3DB-4059-4CDF-9593-D818A11A011D")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Utiliy : IUtility
    {
        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public Int32 x;
            public Int32 y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CURSORINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hCursor;
            public POINT ptScreenPos;
        }

        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(out CURSORINFO pci);

        private const int CURSOR_SHOWING = 0x00000001;


        /// <summary>
        /// 获取鼠标句柄
        /// </summary>
        /// <returns></returns>
        private IntPtr GetCursorHandle()
        {
            CURSORINFO vCurosrInfo;
            vCurosrInfo.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
            GetCursorInfo(out vCurosrInfo);
            return (vCurosrInfo.flags & CURSOR_SHOWING) != CURSOR_SHOWING ? IntPtr.Zero : vCurosrInfo.hCursor;
        }

        /// <summary>
        /// 设置一个超时时间，等待鼠标变回默认指针的状态
        /// 一般用于等待鼠标是转圈的状态
        /// </summary>
        /// <param name="timeoutSecond"></param>
        public void WaitUntilDefaultCursor(int timeoutSecond = 5)
        {
            long startTick = DateTime.Now.Ticks;
            while (true)
            {
                if (Cursors.Default.Handle == GetCursorHandle()) break;

                var elapsedTicks = DateTime.Now.Ticks - startTick;
                if (new TimeSpan(elapsedTicks).TotalSeconds > timeoutSecond)
                {
                    throw new InvalidOperationException($"Wait Until Default Cursor Timeout After {timeoutSecond} seconds");
                }
            }
        }

        /// <summary>
        /// 移动鼠标到指定位置
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="detectScaling"></param>
        public void MoveCursorTo(int x, int y, bool detectScaling = false)
        {
            int xValue = x;
            int yValue = y;
            if (detectScaling) Win32Api.ConvertPoint_LogicalToPhysical(ref xValue, ref yValue);
            Win32Api.Move_Cursor(xValue, yValue);
        }

        /// <summary>
        /// 点击鼠标左键
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="holdSeconds"></param>
        /// <param name="detectScaling"></param>
        public void Click_Left_Mouse(int x, int y, int holdSeconds = 0, bool detectScaling = false)
        {
            int xValue = x;
            int yValue = y;
            if (detectScaling) Win32Api.ConvertPoint_LogicalToPhysical(ref xValue, ref yValue);
            Win32Api.Mouse_Click(xValue, yValue, holdSeconds, "left");
        }

        /// <summary>
        /// 点击鼠标右键
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="holdSeconds"></param>
        /// <param name="detectScaling"></param>
        public void Click_Right_Mouse(int x, int y, int holdSeconds = 0, bool detectScaling = false)
        {
            int xValue = x;
            int yValue = y;
            if (detectScaling) Win32Api.ConvertPoint_LogicalToPhysical(ref xValue, ref yValue);
            Win32Api.Mouse_Click(xValue, yValue, holdSeconds, "right");
        }

        /// <summary>
        /// 用于简单判断两个Object是否相等
        /// </summary>
        /// <param name="obj1"></param>
        /// <param name="obj2"></param>
        /// <returns></returns>
        public bool IsSameObject(object obj1, object obj2)
        {
            string objType1, objType2;
            objType1 = obj1.GetType().FullName;
            objType2 = obj2.GetType().FullName;

            if (!objType1.Equals(objType2)) return false;

            try
            {
                if (obj1 is JabElement)
                {
                    return ((JabElement)(obj1)).Equals(((JabElement)(obj2)));
                }

                return obj1.Equals(obj2);
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
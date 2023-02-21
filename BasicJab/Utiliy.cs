using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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



        private IntPtr GetCursorHandle()
        {
            CURSORINFO vCurosrInfo;
            vCurosrInfo.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
            GetCursorInfo(out vCurosrInfo);
            if ((vCurosrInfo.flags & CURSOR_SHOWING) != CURSOR_SHOWING)
            {
                return IntPtr.Zero;
            }
            else
            {
                return vCurosrInfo.hCursor;
            }
        }

        private Win32Api api;

        /// <summary>
        /// 构造函数
        /// </summary>
        public Utiliy()
        {
            api = new Win32Api();
        }

        /// <summary>
        /// 设置一个超时时间，等待鼠标变回默认指针的状态
        /// 一般用于等待鼠标是转圈的状态
        /// </summary>
        /// <param name="timeout"></param>
        public void WaitUntilDefaultCursor(int timeoutSecond = 5)
        {
            long startTick = DateTime.Now.Ticks;
            while (true)
            {
                if (Cursors.Default.Handle == GetCursorHandle()) break;

                long elapsedTicks = DateTime.Now.Ticks - startTick;
                if (new TimeSpan(elapsedTicks).TotalSeconds > timeoutSecond)
                {
                    throw new InvalidOperationException(String.Format("Wait Until Default Cursor Timeout After {0} seconds", timeoutSecond));
                }
            }
        }

        /// <summary>
        /// 设置鼠标位置
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void MoveCursorTo(int x, int y, bool detect_scaling = false)
        {
            int x_value = x;
            int y_value = y;
            if (detect_scaling) api.ConvertPoint_LogicalToPhysical(ref x_value, ref y_value);
            api.Move_Cursor(x_value, y_value);
        }

        public void Click_Left_Mouse(int x, int y, int holdSeconds = 0, bool detect_scaling = false)
        {
            int x_value = x;
            int y_value = y;
            if (detect_scaling) api.ConvertPoint_LogicalToPhysical(ref x_value, ref y_value);
            api.Mouse_Click(x_value, y_value, holdSeconds, "left");
        }

        public void Click_Right_Mouse(int x, int y, int holdSeconds = 0, bool detect_scaling = false)
        {
            int x_value = x;
            int y_value = y;
            if (detect_scaling) api.ConvertPoint_LogicalToPhysical(ref x_value, ref y_value);
            api.Mouse_Click(x_value, y_value, holdSeconds, "right");
        }

        public bool IsSameObject(object obj1, object obj2)
        {
            string objType1, objType2;
            objType1 = obj1.GetType().FullName;
            objType2 = obj2.GetType().FullName;

            if (!objType1.Equals(objType2)) return false;

            try
            {
                if (obj1 is JabElement){return ((JabElement)(obj1)).Equals(((JabElement)(obj2)));}

                return obj1.Equals(obj2);

            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}

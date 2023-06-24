using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BasicJab.ComInterface
{
    [Guid("F313C0E7-02EA-4D75-B060-CA9672EDF9EE")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUtility
    {
        [Description("等待鼠标变回默认指针状态")]
        void WaitUntilDefaultCursor(int timeoutSecond = 5);
        [Description("移动鼠标到指定的位置\r\n可以设置是否转为当前缩放下的实际坐标")]
        void MoveCursorTo(int x, int y, bool detectScaling = false);
        void Click_Left_Mouse(int x, int y, int holdSeconds = 0, bool detectScaling = false);
        void Click_Right_Mouse(int x, int y, int holdSeconds = 0, bool detectScaling = false);
        [Description("用于判断两个对象是否相等或指向同一内容")]
        bool IsSameObject(object obj1, object obj2);
    }
}

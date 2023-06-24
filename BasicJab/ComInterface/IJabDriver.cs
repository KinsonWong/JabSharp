using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace BasicJab.ComInterface
{
    /// <summary>
    /// JabDriver 对应的Com接口
    /// </summary>
    [Guid("6043992A-07E4-4CF1-81AC-E06C439F63A9")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJabDriver
    {
        IntPtr WHwnd
        {
            [Description("Java应用窗口的句柄")]
            get;
        }
        int JvmId { get; }
        int Pid
        {
            [Description("Java应用的进程ID")]
            get;
        }

        [Description("用指定的标题查找绑定Java窗口")]
        void Init_JabDriver(string title, int timeoutSecond = 10);
        [Description("用Name查找特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement FindElementByName(string strName, bool regexMatch = false);
        [Description("用Role查找特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement FindElementByRole(string strRole, bool regexMatch = false);
        [Description("用Description查找特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement FindElementByDescription(string strDescription, bool regexMatch = false);
        [Description("用State查找特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement FindElementByState(string strState, bool regexMatch = false);
        JabElement FindElementByObjectDepth(int objectDepth);
        JabElement FindElementByIndexInParent(int index);
        JabElement FindElementByChildrenCount(int count);
        [Description("Xpath会遍历所有节点，开销较大，建议少用")]
        JabElement FindElementByXPath(string strXPath);

        [Description("用Name查找一组特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByName(string strName, bool regexMatch = false);
        [Description("用Role查找一组特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByRole(string strRole, bool regexMatch = false);
        [Description("用Description查找一组特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByDescription(string strDescription, bool regexMatch = false);
        [Description("用State查找一组特定的JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByState(string strState, bool regexMatch = false);
        JabElement[] FindElementsByObjectDepth(int objectDepth);
        JabElement[] FindElementsByIndexInParent(int index);
        JabElement[] FindElementsByChildrenCount(int count);
        [Description("Xpath会遍历所有节点，开销较大，建议少用")]
        JabElement[] FindElementsByXPath(string strXPath);

        JabElement WaitUntilElementExists(By by, object value, bool regexMatch = false, int timeoutSecond = 5);

        JabElement GetFocusedElement();
        void PerformKey(SKey shortcutkey);
        void Minimize_Window();
        void Maximize_Window();
        void Set_Window_Size(int width, int height);
        void Set_Window_Position(int x, int y);
        [Description("效果类似于AppActivate")]
        void ActivateWindow();
        [Description("将整个Java窗口截图")]
        void Save_ScreenShot(string savePath);
    }
}

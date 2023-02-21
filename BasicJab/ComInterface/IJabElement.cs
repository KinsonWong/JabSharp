using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BasicJab.ComInterface
{
    [Guid("7BCD5087-584E-48AF-AF9B-AA0CDFFDC0ED")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJabElement
    {
        string name { get; }

        string role
        {
            [Description("注意，程序已将 Role 的空格全部替换为下划线")]
            get;
        }

        string description { get; }
        string states { get; }
        int object_depth { get; }
        int index_in_parent { get; }
        int children_count { get; }
        string text { get; }

        int position_x { get; }
        int position_y { get; }
        int position_width { get; }
        int position_height { get; }

        bool auto_refresh
        {
            [Description("自动刷新默认为False \r\n当设置为True，获取类属性如Name，Role等会先刷新节点")]
            get;
            [Description("自动刷新默认为False \r\n当设置为True，获取类属性如Name，Role等会先刷新节点")]
            set;
        }

        [Description("获取父节点(JabElement) \r\n如果不存在返回nothing")]
        JabElement GetParentElement();
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
        JabElement FindElementByXPath(string strXpath);

        [Description("用Name查找特定的一组JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByName(string strName, bool regexMatch = false);
        [Description("用Role查找特定的一组JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByRole(string strRole, bool regexMatch = false);
        [Description("用Description查找特定的一组JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByDescription(string strDescription, bool regexMatch = false);
        [Description("用State查找特定的一组JabElement，可以设置是否使用正则表达式匹配")]
        JabElement[] FindElementsByState(string strState, bool regexMatch = false);
        JabElement[] FindElementsByObjectDepth(int objectDepth);
        JabElement[] FindElementsByIndexInParent(int index);
        JabElement[] FindElementsByChildrenCount(int count);
        [Description("Xpath会遍历所有节点，开销较大，建议少用")]
        JabElement[] FindElementsByXPath(string strXPath);

        JabElement WaitUntilElementExists(BY byWhat, object value, bool regexMatch = false, int timeoutSecond = 5);

        [Description("返回被选中的第一个Element\r\n 一般用于以下Role:\r\n combo_box\r\n list \r\n page_tab_list")]
        JabElement GetSelectedElement();

        [Description("让 JabElement 获得焦点")]
        bool Request_Focus();
        void Expand();
        [Description("单击行为 \r\n可以设置是否模拟单击 \r\n是否计算缩放后坐标（仅限模拟单击）")]
        void Click(bool isSimulate = false, bool detect_scaling = false);
        void Clear(bool isSimulate = false);
        [Description("可以直接用Java Api赋值 \r\n也可以模拟键盘输入,以及模拟输入是否转义字符\r\n 如果字符串内有 ()~{}^+等字符，建议打开转义")]
        void Send_Text(string value, bool isSimulate = false, bool isEscape = false);
        [Description("直接将文本写入剪切板并粘贴到Element")]
        void Paste_Text(string value);

        [Description("获取Table指定Index的Cell")]
        JabElement Get_Table_Cell(int rowIndex, int ColumnIndex);
        [Description("执行Accessible Action，返回是否执行成功 \r\n 用例:ele.Do_Action(\"Click\")\r\nele.Do_Action(\"Close Window\")")]
        bool Do_Action(string strAction);

        bool IsChecked();
        bool IsEnabled();
        bool IsVisible();
        bool IsShowing();
        bool IsSelected();
        bool IsEditable();
        bool IsActive();
        bool IsFocusable();

    }
}

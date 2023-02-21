using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using WindowsAccessBridgeInterop;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Xml;
using BasicJab.ComInterface;
using BasicJab.Common;
using System.Text.RegularExpressions;

namespace BasicJab
{
    //空间+类名 记录到注册表中，给其他语言创建对象用
    [ProgId("BasicJab.IJabElement")]
    [Guid("B063D2CD-C38C-4BCE-BF40-5DF908EE71F3")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class JabElement : IJabElement
    {
        private AccessBridge objBridge;
        private AccessibleContextNode context_Node;
        private AccessibleActions acc_Action;
        private Win32Api api;
        private IntPtr hwnd;
        private bool _auto_refresh;
        public string elementID;

        /// <summary>
        /// 构造函数 不带参数
        /// </summary>
        /// <param name="node"></param>
        public JabElement()
        {
            api = new Win32Api();
            _auto_refresh = false;
        }

        /// <summary>
        /// 构造函数 重载 带参数
        /// </summary>
        /// <param name="node"></param>
        /// <param name="handle"></param>
        public JabElement(AccessibleContextNode node, IntPtr handle) : this()
        {
            api = new Win32Api();
            hwnd = handle;
            _auto_refresh = false;
            objBridge = node.AccessBridge;
            context_Node = node;
            objBridge.Functions.GetAccessibleActions(context_Node.JvmId, context_Node.AccessibleContextHandle, out acc_Action);
            elementID = Guid.NewGuid().ToString("N");

        }

        /// <summary>
        /// 属性Name
        /// </summary>
        public string name
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().name; }
                catch (Exception) { return string.Empty; }
            }
        }
        /// <summary>
        /// 属性Role
        /// </summary>
        public string role
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().role_en_US.Trim().Replace(" ", "_"); }
                catch (Exception) { return string.Empty; }
            }
        }
        /// <summary>
        /// 属性description
        /// </summary>
        public string description
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().description; }
                catch (Exception) { return string.Empty; }
            }
        }
        /// <summary>
        /// 属性states
        /// </summary>
        public string states
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().states_en_US; }
                catch (Exception) { return string.Empty; }
            }
        }
        /// <summary>
        /// 属性object_depth
        /// </summary>
        public int object_depth
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return objBridge.Functions.GetObjectDepth(context_Node.JvmId, context_Node.AccessibleContextHandle); }
                catch (Exception) { return -1; }
            }
        }
        /// <summary>
        /// 属性index in parent
        /// </summary>
        public int index_in_parent
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().indexInParent; }
                catch (Exception) { return -1; }
            }
        }
        /// <summary>
        /// 属性children count
        /// </summary>
        public int children_count
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().childrenCount; }
                catch (Exception) { return -1; }
            }
        }
        /// <summary>
        /// 属性Text
        /// </summary>
        public string text
        {
            get
            {
                try
                {
                    if (_auto_refresh) { context_Node.Refresh(); }
                    if (context_Node.GetInfo().accessibleText != 0)
                    {
                        AccessibleTextInfo textInfo;
                        if (objBridge.Functions.GetAccessibleTextInfo(context_Node.JvmId, context_Node.AccessibleContextHandle, out textInfo, 0, 0))
                        {
                            var reader = new AccessibleTextReader(context_Node, textInfo.charCount);
                            return reader.ReadToEnd();
                        }
                        else { return string.Empty; }
                    }
                    else { return string.Empty; }
                }
                catch (Exception) { return string.Empty; }
            }
        }
        /// <summary>
        /// 属性坐标x
        /// </summary>
        public int position_x
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().x; }
                catch (Exception) { return -1; }
            }
        }
        /// <summary>
        /// 属性坐标y
        /// </summary>
        public int position_y
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().y; }
                catch (Exception) { return -1; }
            }
        }
        /// <summary>
        /// 属性width
        /// </summary>
        public int position_width
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().width; }
                catch (Exception) { return 0; }
            }
        }
        /// <summary>
        /// 属性height
        /// </summary>
        public int position_height
        {
            get
            {
                try { if (_auto_refresh) { context_Node.Refresh(); } return context_Node.GetInfo().height; }
                catch (Exception) { return 0; }
            }
        }

        /// <summary>
        /// 自动刷新默认为False 
        /// 当设置为True，获取类属性如Name，Role等会先刷新节点
        /// </summary>
        public bool auto_refresh
        {
            get
            {
                return _auto_refresh;
            }
            set
            {
                _auto_refresh = value;
            }
        }

        /// <summary>
        /// 获取父节点 如找不到返回null
        /// </summary>
        /// <returns></returns>
        public JabElement GetParentElement()
        {
            context_Node.Refresh();
            var parent_ac =
                objBridge.Functions.GetAccessibleParentFromContext
                (context_Node.JvmId,
                context_Node.AccessibleContextHandle);

            if (parent_ac.IsNull) return null;
            return new JabElement(new AccessibleContextNode(objBridge, parent_ac), hwnd);
        }

        /// <summary>
        /// 用Name查找特定的一个JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public JabElement FindElementByName(string strName, bool regexMatch = false)
        {
            return Find_Element(BY.NAME, strName, regexMatch);
        }

        /// <summary>
        /// 用Role查找特定的一个JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <returns></returns>
        public JabElement FindElementByRole(string strRole, bool regexMatch = false)
        {
            return Find_Element(BY.ROLE, strRole, regexMatch);
        }

        /// <summary>
        /// 用Description查找特定的一个JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <returns></returns>
        public JabElement FindElementByDescription(string strDescription, bool regexMatch = false)
        {
            return Find_Element(BY.DESCRIPTION, strDescription, regexMatch);
        }

        /// <summary>
        /// 用State查找特定的一个JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <returns></returns>
        public JabElement FindElementByState(string strState, bool regexMatch = false)
        {
            return Find_Element(BY.ROLE, strState, regexMatch);
        }

        /// <summary>
        /// 用Object Depth查找特定的一个JabElement
        /// </summary>
        /// <param name="objectDepth"></param>
        /// <returns></returns>
        public JabElement FindElementByObjectDepth(int objectDepth)
        {
            if (objectDepth < 0) return null;
            return Find_Element(BY.OBJECT_DEPTH, objectDepth);
        }

        /// <summary>
        /// 用index in parent查找特定的一个JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement FindElementByIndexInParent(int index)
        {
            if (index < 0) return null;
            return Find_Element(BY.INDEX_IN_PARENT, index);
        }

        /// <summary>
        /// 用children count查找特定的一个JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement FindElementByChildrenCount(int count)
        {
            if (count < 0) return null;
            return Find_Element(BY.CHILDREN_COUNT, count);
        }

        /// <summary>
        /// 用xpath查找特定的一个Jabelement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXpath"></param>
        /// <returns></returns>
        public JabElement FindElementByXPath(string strXpath)
        {
            var topAC = objBridge.Functions.GetTopLevelObject(context_Node.JvmId, context_Node.AccessibleContextHandle);
            if (topAC.IsNull)
            {
                throw new InvalidOperationException("Can not Find Root Element");
            }
            //找到root节点并初始化Element为对象
            JabElement root = new JabElement(new AccessibleContextNode(objBridge, topAC), hwnd);
            //Debug.WriteLine(root.role);
            Dictionary<string, JabElement> dic = new Dictionary<string, JabElement>();
            XmlDocument doc = new XmlDocument();
            string id = string.Empty;


            //从root节点开始，读取所有子节点，写入到xml中  递归
            XmlElement rootXmlElement = doc.CreateElement(root.role);
            rootXmlElement = Xml_Element_Create(doc, rootXmlElement, ref dic, ref id, root, this);
            doc.AppendChild(rootXmlElement);

            try
            {
                //定位到当前节点
                XmlNode curNode = doc.SelectSingleNode(@"//*[@INTERNAL_GUID='" + id + "']");
                XmlNode targetNode = curNode.SelectSingleNode(@strXpath);

                //返回找到的第一个节点对应的JabElement
                return dic[targetNode.Attributes["INTERNAL_GUID"].Value.ToString()];
            }
            catch (Exception err)
            {
                Debug.WriteLine(err.Message);
                return null;
            }
        }

        //=============================================


        /// <summary>
        /// 用Name查找一组JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByName(string strName, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(BY.NAME, strName, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用Role查找一组JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByRole(string strRole, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(BY.ROLE, strRole, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用Description查找一组JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByDescription(string strDescription, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(BY.DESCRIPTION, strDescription, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用State查找一组JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByState(string strState, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(BY.STATES, strState, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用object depth查找一组JabElement
        /// </summary>
        /// <param name="objectDepth"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByObjectDepth(int objectDepth)
        {
            if (objectDepth < 0) return new List<JabElement>().ToArray();
            List<JabElement> result = Find_Elements(BY.OBJECT_DEPTH, objectDepth);
            return result.ToArray();
        }

        /// <summary>
        /// 用Index in Parent查找一组JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByIndexInParent(int index)
        {
            if (index < 0) return new List<JabElement>().ToArray();
            List<JabElement> result = Find_Elements(BY.INDEX_IN_PARENT, index);
            return result.ToArray();
        }

        /// <summary>
        /// 用Children count查找一组JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByChildrenCount(int count)
        {
            if (count < 0) return new List<JabElement>().ToArray();
            List<JabElement> result = Find_Elements(BY.CHILDREN_COUNT, count);
            return result.ToArray();
        }

        /// <summary>
        /// 用xpath查找一组JabElement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXPath"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByXPath(string strXPath)
        {
            List<JabElement> result = new List<JabElement>();

            var topAC = objBridge.Functions.GetTopLevelObject(context_Node.JvmId, context_Node.AccessibleContextHandle);
            if (topAC.IsNull)
            {
                throw new InvalidOperationException("Can not Find Root Element");
            }
            //找到root节点并初始化Element为对象
            JabElement root = new JabElement(new AccessibleContextNode(objBridge, topAC), hwnd);
            Debug.WriteLine(root.role);
            Dictionary<string, JabElement> dic = new Dictionary<string, JabElement>();
            XmlDocument doc = new XmlDocument();
            string id = string.Empty;


            //从root节点开始，读取所有子节点，写入到xml中  递归
            XmlElement rootXmlElement = doc.CreateElement(root.role);
            rootXmlElement = Xml_Element_Create(doc, rootXmlElement, ref dic, ref id, root, this);
            doc.AppendChild(rootXmlElement);

            try
            {
                //定位到当前节点
                XmlNode curNode = doc.SelectSingleNode(@"//*[@INTERNAL_GUID='" + id + "']");
                XmlNodeList nodes = curNode.SelectNodes(@strXPath);

                //返回找到的所有节点
                foreach (XmlNode n in nodes)
                {
                    result.Add(dic[n.Attributes["INTERNAL_GUID"].Value.ToString()]);
                }
            }
            catch (Exception err)
            {
                Debug.WriteLine(err.Message);
            }
            return result.ToArray();
        }

        /// <summary>
        /// 内部使用 用特定元素查找一个JabElement
        /// </summary>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private JabElement Find_Element(BY byWhat, object value, bool regexMatch = false)
        {
            context_Node.Refresh();
            if (byWhat == BY.XPATH)
            {
                return FindElementByXPath((string)value);
            }
            else
            {
                if (!regexMatch)
                {
                    foreach (JabElement ele in _generate_all_childs(this))
                    {
                        if (_is_element_matched(ele, byWhat, value)) return ele;
                    }
                }
                else
                {
                    foreach (JabElement ele in _generate_all_childs(this))
                    {
                        if (_is_element_regex_matched(ele, byWhat, value)) return ele;
                    }
                }

                return null;

            }
        }

        /// <summary>
        /// 设置一个超时的秒数，一直寻找一个特定的JabElement
        /// </summary>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public JabElement WaitUntilElementExists(BY byWhat, object value, bool regexMatch = false, int timeoutSecond = 5)
        {
            long startTick = DateTime.Now.Ticks;
            JabElement result = null;
            while (true)
            {
                switch (byWhat)
                {
                    case BY.NAME:
                        result = FindElementByName((string)value, regexMatch);
                        break;
                    case BY.ROLE:
                        result = FindElementByRole((string)value, regexMatch);
                        break;
                    case BY.STATES:
                        result = FindElementByState((string)value, regexMatch);
                        break;
                    case BY.DESCRIPTION:
                        result = FindElementByDescription((string)value, regexMatch);
                        break;
                    case BY.OBJECT_DEPTH:
                        result = FindElementByObjectDepth((int)value);
                        break;
                    case BY.INDEX_IN_PARENT:
                        result = FindElementByIndexInParent((int)value);
                        break;
                    case BY.CHILDREN_COUNT:
                        result = FindElementByChildrenCount((int)value);
                        break;
                    case BY.XPATH:
                        result = FindElementByXPath((string)value);
                        break;
                }
                if (result != null) return result;

                long elapsedTicks = DateTime.Now.Ticks - startTick;
                if (new TimeSpan(elapsedTicks).TotalSeconds > timeoutSecond)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 获取被选中的第一个ELement 一般用于 combo box | list | page tab list
        /// </summary>
        /// <returns></returns>
        public JabElement GetSelectedElement()
        {
            try
            {
                var acc = objBridge.Functions.GetAccessibleSelectionFromContext(context_Node.JvmId,context_Node.AccessibleContextHandle, 0);
                return new JabElement(new AccessibleContextNode(objBridge, acc), hwnd);
            }
            catch (Exception)
            {
                return null;
            }
            
        }

        /// <summary>
        /// 内部使用 用特定元素查找一组JabElement
        /// </summary>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private List<JabElement> Find_Elements(BY byWhat, object value, bool regexMatch = false)
        {
            context_Node.Refresh();
            if (byWhat == BY.XPATH)
            {
                return FindElementsByXPath((string)value).ToList();
            }
            else
            {
                List<JabElement> result = new List<JabElement>();

                if (!regexMatch)
                {
                    foreach (JabElement ele in _generate_all_childs(this))
                    {
                        if (_is_element_matched(ele, byWhat, value))
                        {
                            result.Add(ele);
                            continue;
                        }
                    }
                }
                else
                {
                    foreach (JabElement ele in _generate_all_childs(this))
                    {
                        if (_is_element_regex_matched(ele, byWhat, value))
                        {
                            result.Add(ele);
                            continue;
                        }
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// 这是一个递归的方法  将JabElement下的所有子节点写入xml里面
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ParentElement"></param>
        /// <param name="dic"></param>
        /// <param name="id"></param>
        /// <param name="root"></param>
        /// <param name="curEle"></param>
        /// <returns></returns>
        private XmlElement Xml_Element_Create(XmlDocument doc, XmlElement ParentElement,
            ref Dictionary<string, JabElement> dic,
            ref string id, JabElement root, JabElement curEle = null)
        {
            dic.Add(root.elementID, root);

            if (curEle != null)
            {
                if (objBridge.Functions.IsSameObject(context_Node.JvmId, root.context_Node.AccessibleContextHandle, curEle.context_Node.AccessibleContextHandle))
                {
                    id = root.elementID;
                }
            }

            //text 写入 innnerText里面 ，xpath查找时要这样写  //*[text()='123']
            ParentElement.InnerText = root.text;
            //添加其余一般属性
            ParentElement.SetAttribute("INTERNAL_GUID", root.elementID);
            ParentElement.SetAttribute("name", root.name);
            ParentElement.SetAttribute("description", root.description);
            ParentElement.SetAttribute("states", root.states);
            ParentElement.SetAttribute("object_depth", root.object_depth.ToString());
            ParentElement.SetAttribute("index_in_parent", root.index_in_parent.ToString());
            ParentElement.SetAttribute("children_count", root.children_count.ToString());

            //遍历子节点，递归写入xml
            foreach (JabElement ele in _generate_childs_from_element(root))
            {
                XmlElement child = doc.CreateElement(ele.role);

                if (id == string.Empty)
                {
                    child = Xml_Element_Create(doc, child, ref dic, ref id, ele, curEle);
                }
                else
                {
                    child = Xml_Element_Create(doc, child, ref dic, ref id, ele);
                }
                ParentElement.AppendChild(child);
            }
            return ParentElement;
        }

        /// <summary>
        /// 获取当前Element下的所有子元素 迭代器
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>
        private IEnumerable<JabElement> _generate_all_childs(JabElement ele)
        {
            foreach (JabElement _ele in _generate_childs_from_element(ele))
            {
                yield return _ele;
                if (_ele.children_count > 0)
                {
                    foreach (JabElement child in _generate_all_childs(_ele))
                    {
                        yield return child;
                    }
                }
            }
        }

        /// <summary>
        /// 获取当前Element下的所有子元素 迭代器
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>
        private IEnumerable<JabElement> _generate_childs_from_element(JabElement ele)
        {
            foreach (AccessibleContextNode node in ele.context_Node.GetChildren())
            {
                yield return new JabElement(node, hwnd);
            }

        }

        /// <summary>
        /// 判断JabElement是否匹配特定的元素
        /// </summary>
        /// <param name="ele"></param>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _is_element_matched(JabElement ele, BY byWhat, object value)
        {
            switch (byWhat)
            {
                case BY.NAME:
                    return ele.name == (string)value;
                case BY.DESCRIPTION:
                    return ele.description == (string)value;
                case BY.ROLE:
                    return ele.role == (string)value;
                case BY.STATES:
                    return ele.states == (string)value;
                case BY.OBJECT_DEPTH:
                    return ele.object_depth == (int)value;
                case BY.CHILDREN_COUNT:
                    return ele.children_count == (int)value;
                case BY.INDEX_IN_PARENT:
                    return ele.index_in_parent == (int)value;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 用正则表达式模糊匹配
        /// </summary>
        /// <param name="ele"></param>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _is_element_regex_matched(JabElement ele, BY byWhat, object value)
        {
            switch (byWhat)
            {
                case BY.NAME:
                    return new Regex((string)@value).IsMatch(ele.name);
                case BY.DESCRIPTION:
                    return new Regex((string)@value).IsMatch(ele.description);
                case BY.ROLE:
                    return new Regex((string)@value).IsMatch(ele.role);
                case BY.STATES:
                    return new Regex((string)@value).IsMatch(ele.name);

                //以下非字符串属性不能用正则匹配
                case BY.OBJECT_DEPTH:
                    return ele.object_depth == (int)value;
                case BY.CHILDREN_COUNT:
                    return ele.children_count == (int)value;
                case BY.INDEX_IN_PARENT:
                    return ele.index_in_parent == (int)value;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 让当前Element获得焦点
        /// </summary>
        /// <returns></returns>
        public bool Request_Focus()
        {
            return objBridge.Functions.RequestFocus(context_Node.JvmId, context_Node.AccessibleContextHandle);
        }

        /// <summary>
        /// /展开操作 
        /// </summary>
        /// <param name="isSimulate"></param>
        public void Expand()
        {
            if (!states.Contains("expandable"))
            {
                throw new InvalidOperationException("Element dosen't support \"Expand\" Action");
            }
            Do_Accessible_Action("toggleexpand");
        }

        /// <summary>
        /// 点击Element 方法，可以设置是否使用模拟点击
        /// </summary>
        public void Click(bool isSimulate = false, bool detect_scaling = false)
        {
            if (isSimulate)
            {
                api.SetTopWindow(hwnd);
                Request_Focus();
                if (position_height == 0 || position_width == 0)
                {
                    throw new InvalidOperationException("Element height or width can not equal to Zero!");
                }

                int r_x = (int)Math.Round(position_x + (double)position_width / 2);
                int r_y = (int)Math.Round(position_y + (double)position_height / 2);
                if (detect_scaling) api.ConvertPoint_LogicalToPhysical(ref r_x, ref r_y);
                api.Mouse_Click(r_x, r_y);
            }
            else
            {
                if (!Do_Accessible_Action("click"))
                {
                    if (!Do_Accessible_Action("单击"))
                    {
                        throw new InvalidOperationException("This element doesn't support \"Click\" action!");
                    }
                }
            }
        }

        /// <summary>
        /// 清空当前Element文本，可以设置是否模拟操作
        /// </summary>
        /// <param name="isSimulate"></param>
        public void Clear(bool isSimulate = false)
        {
            if (isSimulate)
            {
                api.SetTopWindow(hwnd);
                Request_Focus();
                if (text != string.Empty && context_Node.GetInfo().accessibleText != 0)
                {
                    //通过以下操作全选文本
                    SendKeys.SendWait("{HOME}");
                    SendKeys.SendWait("^+{END}");  //ctrl+shift+end

                    for (int i = 0; i < 50; i++)   //保险起见 循环50次 输入 退格 正常一次就够
                    {
                        SendKeys.SendWait("{BACKSPACE}");
                    }
                }
            }
            else //尝试用Java Bridge API去清除文本
            {
                Send_Text(string.Empty, false);
            }

        }

        /// <summary>
        /// 对指定Element进行输入文本操作，可以设置是否模拟键盘输入
        /// </summary>
        /// <param name="value"></param>
        /// <param name="isSimulate"></param>
        public void Send_Text(string value, bool isSimulate = false, bool isEscape = false)
        {
            if (isSimulate)
            {
                api.SetTopWindow(hwnd);
                Request_Focus();
                if (isEscape)
                {
                    foreach (char c in value)
                    {
                        if (c.ToString() == "(")
                            SendKeys.SendWait("{(}");
                        else if (c.ToString() == ")")
                            SendKeys.SendWait("{)}");
                        else if (c.ToString() == "^")
                            SendKeys.SendWait("{^}");
                        else if (c.ToString() == "+")
                            SendKeys.SendWait("{+}");
                        else if (c.ToString() == "%")
                            SendKeys.SendWait("{%}");
                        else if (c.ToString() == "~")
                            SendKeys.SendWait("{~}");
                        else if (c.ToString() == "{")
                            SendKeys.SendWait("{{}");
                        else if (c.ToString() == "}")
                            SendKeys.SendWait("{}}");
                        else
                            SendKeys.SendWait(c.ToString());
                    }
                }
                else { SendKeys.SendWait(value); }

            }
            else
            {
                bool result = objBridge.Functions.SetTextContents(context_Node.JvmId, context_Node.AccessibleContextHandle, value);
                if (!result)
                {
                    throw new InvalidOperationException("JAB Function 'SetTextContents' Failed!\nTry set parameter 'isSimulate' With 'True'!");
                }
            }
        }

        /// <summary>
        /// 调Win API将文本写入剪切板，再粘贴，应该比模拟输入快一点
        /// </summary>
        /// <param name="value"></param>
        public void Paste_Text(string value)
        {
            Clipboard_Util util = new Clipboard_Util();
            util.SetText(value);
            api.SetTopWindow(hwnd);
            Request_Focus();
            SendKeys.SendWait("^{v}");    //ctrl+v
            Delay(200);
        }
        
        /// <summary>
        /// 获取Table 指定index的cell
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        public JabElement Get_Table_Cell(int rowIndex, int ColumnIndex)
        {
            if (!role.ToLower().Equals("table"))
            {
                throw new InvalidOperationException("Current JabElement is not a Table!");
            }

            AccessibleTableInfo Table_Info = new AccessibleTableInfo();
            if(!objBridge.Functions.GetAccessibleTableInfo(context_Node.JvmId, context_Node.AccessibleContextHandle, out Table_Info))
            {
                throw new InvalidOperationException("can not get table info!");
            }

            AccessibleTableCellInfo Cell_Info = new AccessibleTableCellInfo();
            if(!objBridge.Functions.GetAccessibleTableCellInfo(context_Node.JvmId,Table_Info.accessibleTable,rowIndex,ColumnIndex, out Cell_Info))
            {
                throw new InvalidOperationException("can not get table cell info!");
            }
            return new JabElement(new AccessibleContextNode(objBridge,Cell_Info.accessibleContext), hwnd);
        }

        public bool IsChecked()
        {
            context_Node.Refresh();
            return states.Contains("checked");
        }

        public bool IsEnabled()
        {
            context_Node.Refresh();
            return states.Contains("enabled");
        }

        public bool IsVisible()
        {
            context_Node.Refresh();
            return states.Contains("visible");
        }

        public bool IsShowing()
        {
            context_Node.Refresh();
            return states.Contains("showing");
        }

        public bool IsSelected()
        {
            context_Node.Refresh();
            return states.Contains("selected");
        }

        public bool IsEditable()
        {
            context_Node.Refresh();
            return states.Contains("editable");
        }

        public bool IsActive()
        {
            context_Node.Refresh();
            return states.Contains("active");
        }

        public bool IsFocusable()
        {
            context_Node.Refresh();
            return states.Contains("focusable");
        }

        /// <summary>
        /// 外部调用 手动指定一种Action
        /// </summary>
        /// <param name="strAction"></param>
        /// <returns></returns>
        public bool Do_Action(string strAction)
        {
            return Do_Accessible_Action(strAction);
        }

        /// <summary>
        /// 传入具体的action字符串，返回是否能通过JAB执行action
        /// </summary>
        /// <param name="action"></param>
        /// <returns></returns>
        private bool Do_Accessible_Action(string action)
        {
            if (acc_Action.actionsCount <= 0) return false;

            AccessibleActionInfo[] infoArr = acc_Action.actionInfo;
            int actQueryCount =
                (from info in infoArr
                 where info.name.ToLower() == action.ToLower()
                 select info).Count();

            if (actQueryCount <= 0) return false;

            AccessibleActionInfo targetActionInfo =
                (from info in infoArr
                 where info.name.ToLower() == action.ToLower()
                 select info).First();

            AccessibleActionsToDo todo = new AccessibleActionsToDo();
            todo.actions = new AccessibleActionInfo[32];
            todo.actions[0] = targetActionInfo;
            todo.actionsCount = 1;

            int hrResult;
            return objBridge.Functions.DoAccessibleActions(context_Node.JvmId, context_Node.AccessibleContextHandle, ref todo, out hrResult);
        }

        /// <summary>
        /// 内部方法 另一种方式实现Sleep
        /// </summary>
        private void Delay(int milliseconds)
        {
            long startTick = DateTime.Now.Ticks;
            while (true)
            {
                if (new TimeSpan(DateTime.Now.Ticks - startTick).TotalMilliseconds > milliseconds) return;
            }
        }

        /// <summary>
        /// 判断两个Element是否指向同一个对象
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>
        public bool Equals(JabElement ele)
        {
            if (objBridge.Functions.IsSameObject(context_Node.JvmId, context_Node.AccessibleContextHandle, ele.context_Node.AccessibleContextHandle))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}

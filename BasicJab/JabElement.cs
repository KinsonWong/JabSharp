using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using WindowsAccessBridgeInterop;
using System.Diagnostics;
using System.Windows.Forms;
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
        private AccessBridge _objBridge;
        private AccessibleContextNode _contextNode;
        private AccessibleActions _accAction;
        private IntPtr _hwnd;
        private bool _autoRefresh;
        public string elementId;

        /// <summary>
        /// 构造函数 不带参数
        /// </summary>
        public JabElement()
        {
            _autoRefresh = false;
        }

        /// <summary>
        /// 构造函数 重载 带参数
        /// </summary>
        /// <param name="node"></param>
        /// <param name="handle"></param>
        public JabElement(AccessibleContextNode node, IntPtr handle) : this()
        {
            _hwnd = handle;
            _autoRefresh = false;
            _objBridge = node.AccessBridge;
            _contextNode = node;
            _objBridge.Functions.GetAccessibleActions(_contextNode.JvmId, _contextNode.AccessibleContextHandle,
                out _accAction);
            elementId = Guid.NewGuid().ToString("N");
        }

        /// <summary>
        /// 属性Name
        /// </summary>
        public string Name
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().name;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 属性Role
        /// 注意 -> 为了实现Xpath定位，Role中的空格被全部替换为下划线
        /// </summary>
        public string Role
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().role_en_US.Trim().Replace(" ", "_");
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 属性description
        /// </summary>
        public string Description
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().description;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 属性states
        /// </summary>
        public string States
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().states_en_US;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 属性object_depth
        /// </summary>
        public int ObjectDepth
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _objBridge.Functions.GetObjectDepth(_contextNode.JvmId,
                        _contextNode.AccessibleContextHandle);
                }
                catch (Exception)
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// 属性index in parent
        /// </summary>
        public int IndexInParent
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().indexInParent;
                }
                catch (Exception)
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// 属性children count
        /// </summary>
        public int ChildrenCount
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().childrenCount;
                }
                catch (Exception)
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// 属性Text
        /// </summary>
        public string Text
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    if (_contextNode.GetInfo().accessibleText != 0)
                    {
                        AccessibleTextInfo textInfo;
                        if (_objBridge.Functions.GetAccessibleTextInfo(_contextNode.JvmId,
                                _contextNode.AccessibleContextHandle, out textInfo, 0, 0))
                        {
                            var reader = new AccessibleTextReader(_contextNode, textInfo.charCount);
                            return reader.ReadToEnd();
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 属性坐标x
        /// </summary>
        public int PositionX
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().x;
                }
                catch (Exception)
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// 属性坐标y
        /// </summary>
        public int PositionY
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().y;
                }
                catch (Exception)
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// 属性width
        /// </summary>
        public int PositionWidth
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().width;
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// 属性height
        /// </summary>
        public int PositionHeight
        {
            get
            {
                try
                {
                    if (_autoRefresh)
                    {
                        _contextNode.Refresh();
                    }

                    return _contextNode.GetInfo().height;
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// 自动刷新默认为False 
        /// 当设置为True，获取类属性如Name，Role等会先刷新节点
        /// </summary>
        public bool AutoRefresh
        {
            get => _autoRefresh;
            set => _autoRefresh = value;
        }

        /// <summary>
        /// 获取父节点 如找不到返回null
        /// </summary>
        /// <returns></returns>
        public JabElement GetParentElement()
        {
            _contextNode.Refresh();
            var parentAc =
                _objBridge.Functions.GetAccessibleParentFromContext
                (_contextNode.JvmId,
                    _contextNode.AccessibleContextHandle);

            if (parentAc.IsNull) return null;
            return new JabElement(new AccessibleContextNode(_objBridge, parentAc), _hwnd);
        }

        /// <summary>
        /// 用Name查找特定的一个JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByName(string strName, bool regexMatch = false)
        {
            return Find_Element(By.NAME, strName, regexMatch);
        }

        /// <summary>
        /// 用Role查找特定的一个JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByRole(string strRole, bool regexMatch = false)
        {
            return Find_Element(By.ROLE, strRole, regexMatch);
        }

        /// <summary>
        /// 用Description查找特定的一个JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByDescription(string strDescription, bool regexMatch = false)
        {
            return Find_Element(By.DESCRIPTION, strDescription, regexMatch);
        }

        /// <summary>
        /// 用State查找特定的一个JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByState(string strState, bool regexMatch = false)
        {
            return Find_Element(By.ROLE, strState, regexMatch);
        }

        /// <summary>
        /// 用Object Depth查找特定的一个JabElement
        /// </summary>
        /// <param name="objectDepth"></param>
        /// <returns></returns>
        public JabElement FindElementByObjectDepth(int objectDepth)
        {
            if (objectDepth < 0) return null;
            return Find_Element(By.OBJECT_DEPTH, objectDepth);
        }

        /// <summary>
        /// 用index in parent查找特定的一个JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement FindElementByIndexInParent(int index)
        {
            if (index < 0) return null;
            return Find_Element(By.INDEX_IN_PARENT, index);
        }

        /// <summary>
        /// 用children count查找特定的一个JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement FindElementByChildrenCount(int count)
        {
            if (count < 0) return null;
            return Find_Element(By.CHILDREN_COUNT, count);
        }

        /// <summary>
        /// 用xpath查找特定的一个Jabelement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXpath"></param>
        /// <returns></returns>
        public JabElement FindElementByXPath(string strXpath)
        {
            var topAc = _objBridge.Functions.GetTopLevelObject(_contextNode.JvmId,
                _contextNode.AccessibleContextHandle);
            if (topAc.IsNull)
            {
                throw new InvalidOperationException("Can not Find Root Element");
            }

            //找到root节点并初始化Element为对象
            JabElement root = new JabElement(new AccessibleContextNode(_objBridge, topAc), _hwnd);
            //Debug.WriteLine(root.role);
            Dictionary<string, JabElement> dic = new Dictionary<string, JabElement>();
            XmlDocument doc = new XmlDocument();
            string id = string.Empty;


            //从root节点开始，读取所有子节点，写入到xml中  递归
            XmlElement rootXmlElement = doc.CreateElement(root.Role);
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
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByName(string strName, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(By.NAME, strName, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用Role查找一组JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByRole(string strRole, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(By.ROLE, strRole, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用Description查找一组JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByDescription(string strDescription, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(By.DESCRIPTION, strDescription, regexMatch);
            return result.ToArray();
        }

        /// <summary>
        /// 用State查找一组JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByState(string strState, bool regexMatch = false)
        {
            List<JabElement> result = Find_Elements(By.STATES, strState, regexMatch);
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
            List<JabElement> result = Find_Elements(By.OBJECT_DEPTH, objectDepth);
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
            List<JabElement> result = Find_Elements(By.INDEX_IN_PARENT, index);
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
            List<JabElement> result = Find_Elements(By.CHILDREN_COUNT, count);
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

            var topAc = _objBridge.Functions.GetTopLevelObject(_contextNode.JvmId,
                _contextNode.AccessibleContextHandle);
            if (topAc.IsNull)
            {
                throw new InvalidOperationException("Can not Find Root Element");
            }

            //找到root节点并初始化Element为对象
            JabElement root = new JabElement(new AccessibleContextNode(_objBridge, topAc), _hwnd);
            Debug.WriteLine(root.Role);
            Dictionary<string, JabElement> dic = new Dictionary<string, JabElement>();
            XmlDocument doc = new XmlDocument();
            string id = string.Empty;


            //从root节点开始，读取所有子节点，写入到xml中  递归
            XmlElement rootXmlElement = doc.CreateElement(root.Role);
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
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        private JabElement Find_Element(By byWhat, object value, bool regexMatch = false)
        {
            _contextNode.Refresh();
            if (byWhat == By.XPATH)
            {
                return FindElementByXPath((string)value);
            }
            else
            {
                if (!regexMatch)
                {
                    foreach (JabElement ele in _generate_all_children(this))
                    {
                        if (_is_element_matched(ele, byWhat, value)) return ele;
                    }
                }
                else
                {
                    foreach (JabElement ele in _generate_all_children(this))
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
        /// <param name="regexMatch"></param>
        /// <param name="timeoutSecond"></param>
        /// <returns></returns>
        public JabElement WaitUntilElementExists(By byWhat, object value, bool regexMatch = false,
            int timeoutSecond = 5)
        {
            long startTick = DateTime.Now.Ticks;
            JabElement result = null;
            while (true)
            {
                switch (byWhat)
                {
                    case By.NAME:
                        result = FindElementByName((string)value, regexMatch);
                        break;
                    case By.ROLE:
                        result = FindElementByRole((string)value, regexMatch);
                        break;
                    case By.STATES:
                        result = FindElementByState((string)value, regexMatch);
                        break;
                    case By.DESCRIPTION:
                        result = FindElementByDescription((string)value, regexMatch);
                        break;
                    case By.OBJECT_DEPTH:
                        result = FindElementByObjectDepth((int)value);
                        break;
                    case By.INDEX_IN_PARENT:
                        result = FindElementByIndexInParent((int)value);
                        break;
                    case By.CHILDREN_COUNT:
                        result = FindElementByChildrenCount((int)value);
                        break;
                    case By.XPATH:
                        result = FindElementByXPath((string)value);
                        break;
                }

                if (result != null) return result;

                var elapsedTicks = DateTime.Now.Ticks - startTick;
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
                var acc = _objBridge.Functions.GetAccessibleSelectionFromContext(_contextNode.JvmId,
                    _contextNode.AccessibleContextHandle, 0);
                return new JabElement(new AccessibleContextNode(_objBridge, acc), _hwnd);
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
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        private List<JabElement> Find_Elements(By byWhat, object value, bool regexMatch = false)
        {
            _contextNode.Refresh();
            if (byWhat == By.XPATH)
            {
                return FindElementsByXPath((string)value).ToList();
            }
            else
            {
                List<JabElement> result = new List<JabElement>();

                if (!regexMatch)
                {
                    foreach (JabElement ele in _generate_all_children(this))
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
                    foreach (JabElement ele in _generate_all_children(this))
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
        /// <param name="parentElement"></param>
        /// <param name="dic"></param>
        /// <param name="id"></param>
        /// <param name="root"></param>
        /// <param name="curEle"></param>
        /// <returns></returns>
        private XmlElement Xml_Element_Create(XmlDocument doc, XmlElement parentElement,
            ref Dictionary<string, JabElement> dic,
            ref string id, JabElement root, JabElement curEle = null)
        {
            dic.Add(root.elementId, root);

            if (curEle != null)
            {
                if (_objBridge.Functions.IsSameObject(_contextNode.JvmId, root._contextNode.AccessibleContextHandle,
                        curEle._contextNode.AccessibleContextHandle))
                {
                    id = root.elementId;
                }
            }

            //text 写入 innerText里面 ，xpath查找时要这样写  //*[text()='123']
            parentElement.InnerText = root.Text;
            //添加其余一般属性
            parentElement.SetAttribute("INTERNAL_GUID", root.elementId);
            parentElement.SetAttribute("name", root.Name);
            parentElement.SetAttribute("description", root.Description);
            parentElement.SetAttribute("states", root.States);
            parentElement.SetAttribute("object_depth", root.ObjectDepth.ToString());
            parentElement.SetAttribute("index_in_parent", root.IndexInParent.ToString());
            parentElement.SetAttribute("children_count", root.ChildrenCount.ToString());

            //遍历子节点，递归写入xml
            foreach (JabElement ele in _generate_children_from_element(root))
            {
                XmlElement child = doc.CreateElement(ele.Role);

                if (id == string.Empty)
                {
                    child = Xml_Element_Create(doc, child, ref dic, ref id, ele, curEle);
                }
                else
                {
                    child = Xml_Element_Create(doc, child, ref dic, ref id, ele);
                }

                parentElement.AppendChild(child);
            }

            return parentElement;
        }

        /// <summary>
        /// 获取当前Element下的所有子元素 迭代器
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>
        private IEnumerable<JabElement> _generate_all_children(JabElement ele)
        {
            foreach (JabElement _ele in _generate_children_from_element(ele))
            {
                yield return _ele;
                if (_ele.ChildrenCount > 0)
                {
                    foreach (JabElement child in _generate_all_children(_ele))
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
        private IEnumerable<JabElement> _generate_children_from_element(JabElement ele)
        {
            foreach (var accessibleNode in ele._contextNode.GetChildren())
            {
                var node = (AccessibleContextNode)accessibleNode;
                yield return new JabElement(node, _hwnd);
            }
        }

        /// <summary>
        /// 判断JabElement是否匹配特定的元素
        /// </summary>
        /// <param name="ele"></param>
        /// <param name="byWhat"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _is_element_matched(JabElement ele, By byWhat, object value)
        {
            switch (byWhat)
            {
                case By.NAME:
                    return ele.Name == (string)value;
                case By.DESCRIPTION:
                    return ele.Description == (string)value;
                case By.ROLE:
                    return ele.Role == (string)value;
                case By.STATES:
                    return ele.States == (string)value;
                case By.OBJECT_DEPTH:
                    return ele.ObjectDepth == (int)value;
                case By.CHILDREN_COUNT:
                    return ele.ChildrenCount == (int)value;
                case By.INDEX_IN_PARENT:
                    return ele.IndexInParent == (int)value;
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
        private bool _is_element_regex_matched(JabElement ele, By byWhat, object value)
        {
            switch (byWhat)
            {
                case By.NAME:
                    return new Regex((string)@value).IsMatch(ele.Name);
                case By.DESCRIPTION:
                    return new Regex((string)@value).IsMatch(ele.Description);
                case By.ROLE:
                    return new Regex((string)@value).IsMatch(ele.Role);
                case By.STATES:
                    return new Regex((string)@value).IsMatch(ele.Name);

                //以下非字符串属性不能用正则匹配
                case By.OBJECT_DEPTH:
                    return ele.ObjectDepth == (int)value;
                case By.CHILDREN_COUNT:
                    return ele.ChildrenCount == (int)value;
                case By.INDEX_IN_PARENT:
                    return ele.IndexInParent == (int)value;
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
            return _objBridge.Functions.RequestFocus(_contextNode.JvmId, _contextNode.AccessibleContextHandle);
        }

        /// <summary>
        /// /展开操作 
        /// </summary>
        public void Expand()
        {
            if (!States.Contains("expandable"))
            {
                throw new InvalidOperationException("Element doesn't support \"Expand\" Action");
            }

            Do_Accessible_Action("toggleexpand");
        }

        /// <summary>
        /// 点击Element 方法，可以设置是否使用模拟点击
        /// </summary>
        /// <param name="isSimulate"></param>
        /// <param name="detectScaling"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Click(bool isSimulate = false, bool detectScaling = false)
        {
            if (isSimulate)
            {
                Win32Api.SetTopWindow(_hwnd);
                Request_Focus();
                if (PositionHeight == 0 || PositionWidth == 0)
                {
                    throw new InvalidOperationException("Element height or width can not equal to Zero!");
                }

                int r_x = (int)Math.Round(PositionX + (double)PositionWidth / 2);
                int r_y = (int)Math.Round(PositionY + (double)PositionHeight / 2);
                if (detectScaling) Win32Api.ConvertPoint_LogicalToPhysical(ref r_x, ref r_y);
                Win32Api.Mouse_Click(r_x, r_y);
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
                Win32Api.SetTopWindow(_hwnd);
                Request_Focus();
                if (Text != string.Empty && _contextNode.GetInfo().accessibleText != 0)
                {
                    //通过以下操作全选文本
                    SendKeys.SendWait("{HOME}");
                    SendKeys.SendWait("^+{END}"); //ctrl+shift+end

                    for (int i = 0; i < 50; i++) //保险起见 循环50次 输入 退格 正常一次就够
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
        /// <param name="isEscape"></param>
        public void Send_Text(string value, bool isSimulate = false, bool isEscape = false)
        {
            if (isSimulate)
            {
                Win32Api.SetTopWindow(_hwnd);
                Request_Focus();
                if (isEscape)
                {
                    foreach (char c in value)
                    {
                        switch (c.ToString())
                        {
                            case "(":
                                SendKeys.SendWait("{(}");
                                break;
                            case ")":
                                SendKeys.SendWait("{)}");
                                break;
                            case "^":
                                SendKeys.SendWait("{^}");
                                break;
                            case "+":
                                SendKeys.SendWait("{+}");
                                break;
                            case "%":
                                SendKeys.SendWait("{%}");
                                break;
                            case "~":
                                SendKeys.SendWait("{~}");
                                break;
                            case "{":
                                SendKeys.SendWait("{{}");
                                break;
                            case "}":
                                SendKeys.SendWait("{}}");
                                break;
                            default:
                                SendKeys.SendWait(c.ToString());
                                break;
                        }
                    }
                }
                else
                {
                    SendKeys.SendWait(value);
                }
            }
            else
            {
                bool result = _objBridge.Functions.SetTextContents(_contextNode.JvmId,
                    _contextNode.AccessibleContextHandle, value);
                if (!result)
                {
                    throw new InvalidOperationException(
                        "JAB Function 'SetTextContents' Failed!\nTry set parameter 'isSimulate' With 'True'!");
                }
            }
        }

        /// <summary>
        /// 调Win API将文本写入剪切板，再粘贴，应该比模拟输入快一点
        /// </summary>
        /// <param name="value"></param>
        public void Paste_Text(string value)
        {
            Clipboard_Util.SetText(value);
            Win32Api.SetTopWindow(_hwnd);
            Request_Focus();
            SendKeys.SendWait("^{v}"); //ctrl+v
            Delay(200);
        }

        /// <summary>
        /// 获取Table 指定index的cell
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public JabElement Get_Table_Cell(int rowIndex, int columnIndex)
        {
            if (!Role.ToLower().Equals("table"))
            {
                throw new InvalidOperationException("Current JabElement is not a Table!");
            }

            AccessibleTableInfo tableInfo = new AccessibleTableInfo();
            if (!_objBridge.Functions.GetAccessibleTableInfo(_contextNode.JvmId, _contextNode.AccessibleContextHandle,
                    out tableInfo))
            {
                throw new InvalidOperationException("can not get table info!");
            }

            AccessibleTableCellInfo cellInfo = new AccessibleTableCellInfo();
            if (!_objBridge.Functions.GetAccessibleTableCellInfo(_contextNode.JvmId, tableInfo.accessibleTable,
                    rowIndex, columnIndex, out cellInfo))
            {
                throw new InvalidOperationException("can not get table cell info!");
            }

            return new JabElement(new AccessibleContextNode(_objBridge, cellInfo.accessibleContext), _hwnd);
        }

        public bool IsChecked()
        {
            _contextNode.Refresh();
            return States.Contains("checked");
        }

        public bool IsEnabled()
        {
            _contextNode.Refresh();
            return States.Contains("enabled");
        }

        public bool IsVisible()
        {
            _contextNode.Refresh();
            return States.Contains("visible");
        }

        public bool IsShowing()
        {
            _contextNode.Refresh();
            return States.Contains("showing");
        }

        public bool IsSelected()
        {
            _contextNode.Refresh();
            return States.Contains("selected");
        }

        public bool IsEditable()
        {
            _contextNode.Refresh();
            return States.Contains("editable");
        }

        public bool IsActive()
        {
            _contextNode.Refresh();
            return States.Contains("active");
        }

        public bool IsFocusable()
        {
            _contextNode.Refresh();
            return States.Contains("focusable");
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
            if (_accAction.actionsCount <= 0) return false;

            AccessibleActionInfo[] infoArr = _accAction.actionInfo;
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
            return _objBridge.Functions.DoAccessibleActions(_contextNode.JvmId, _contextNode.AccessibleContextHandle,
                ref todo, out hrResult);
        }

        /// <summary>
        /// 内部方法 另一种方式实现Sleep
        /// </summary>
        private void Delay(int milliseconds)
        {
            var startTick = DateTime.Now.Ticks;
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
            return _objBridge.Functions.IsSameObject(_contextNode.JvmId, _contextNode.AccessibleContextHandle,
                ele._contextNode.AccessibleContextHandle);
        }
    }
}
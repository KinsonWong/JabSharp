using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using BasicJab.ComInterface;
using BasicJab.Common;
using System.Text.RegularExpressions;
using WindowsAccessBridgeInterop;

namespace BasicJab
{
    //空间+类名 记录到注册表中，给其他语言创建对象用
    [ProgId("BasicJab.IJabDriver")]
    [Guid("473487BA-66D1-49E7-B291-D72B047BF600")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class JabDriver : IJabDriver
    {
        private AccessBridge _objBridge;
        private IntPtr _hwnd;
        private AccessibleWindow _javaWindow;
        private JabElement _rootElement;
        private STAThread _messageThread;

        private Dictionary<SKey, string> _shortcutDic;

        public IntPtr WHwnd => _hwnd;

        public int JvmId => _javaWindow.JvmId;

        public int Pid => Win32Api.GetPidFromHwnd(_hwnd);


        /// <summary>
        /// 构造函数
        /// </summary>
        public JabDriver()
        {
            //初始化Key字典
            Init_ShortcutDic();

            _javaWindow = null;
            _hwnd = IntPtr.Zero;
            //由于Java Access Bridge不能在非UI线程上运行（没有消息泵）
            //所以需要单独开一个线程(STA Thread)
            _messageThread = new STAThread();
            _objBridge = new AccessBridge();
            _objBridge.Initilized += (sender1, args) => { Debug.WriteLine("Initialize Bridge Successfully"); }; //Bridge初始化成功就打印这句话
            _messageThread.Invoke(new Action(() => { _objBridge.Initialize(); }), new object[] { });
        }

        /// <summary>
        /// 析构函数
        /// </summary>
        ~JabDriver()
        {
            try
            {
                _messageThread.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 外部调用 初始化JabDriver对象
        /// </summary>
        /// <param name="title"></param>
        /// <param name="timeoutSecond"></param>
        public void Init_JabDriver(string title, int timeoutSecond = 10)
        {
            var startTick = DateTime.Now.Ticks;
            while (true)
            {
                foreach (KeyValuePair<IntPtr, string> window in Win32Api.GetAllOpenWindows())
                {
                    if (window.Value != title) continue;
                    _hwnd = window.Key;
                    break;
                }

                //hwnd = api.FindWindowByCaption(title);
                Func<AccessBridge, IntPtr, AccessibleWindow> f1 = GetJavaWindow;
                _javaWindow = (AccessibleWindow)_messageThread.Invoke(f1, new object[] { _objBridge, _hwnd });
                if (_javaWindow != null) break;

                var elapsedTicks = DateTime.Now.Ticks - startTick;
                if (new TimeSpan(elapsedTicks).TotalSeconds > timeoutSecond)
                {
                    _messageThread.Dispose();
                    throw new InvalidOperationException(
                        String.Format("Can't find java window by title '{0}' in '{1}' seconds", title, timeoutSecond));
                }
            }

            Win32Api.SetTopWindow(_hwnd); //java窗口置顶
            _rootElement = new JabElement(_javaWindow, _hwnd); //根节点赋值
        }

        /// <summary>
        /// 寻找指定Name的JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByName(string strName, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                return strName == _rootElement.Name ? _rootElement : _rootElement.FindElementByName(strName);
            }
            else
            {
                return new Regex(@strName).IsMatch(_rootElement.Name) ? _rootElement : _rootElement.FindElementByName(strName, true);
            }
        }

        /// <summary>
        /// 寻找指定Role的JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByRole(string strRole, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                return strRole == _rootElement.Role ? _rootElement : _rootElement.FindElementByRole(strRole);
            }
            else
            {
                return new Regex(@strRole).IsMatch(_rootElement.Role) ? _rootElement : _rootElement.FindElementByRole(strRole, regexMatch);
            }
        }

        /// <summary>
        /// 寻找指定Description的JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByDescription(string strDescription, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                return strDescription == _rootElement.Description ? _rootElement : _rootElement.FindElementByDescription(strDescription);
            }
            else
            {
                return new Regex(@strDescription).IsMatch(_rootElement.Description)
                    ? _rootElement
                    : _rootElement.FindElementByDescription(strDescription, regexMatch);
            }
        }

        /// <summary>
        /// 用State查找特定一个JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement FindElementByState(string strState, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                return strState == _rootElement.States ? _rootElement : _rootElement.FindElementByState(strState);
            }
            else
            {
                return new Regex(@strState).IsMatch(_rootElement.States)
                    ? _rootElement
                    : _rootElement.FindElementByState(strState, regexMatch);
            }
        }

        /// <summary>
        /// 用Object Depth查找特定一个JabElement
        /// </summary>
        /// <param name="objectDepth"></param>
        /// <returns></returns>
        public JabElement FindElementByObjectDepth(int objectDepth)
        {
            if (objectDepth < 0) return null;
            return objectDepth == _rootElement.ObjectDepth ? _rootElement : _rootElement.FindElementByObjectDepth(objectDepth);
        }

        /// <summary>
        /// 用 index in parent 查找特定的一个JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement FindElementByIndexInParent(int index)
        {
            if (index < 0) return null;
            return index == _rootElement.IndexInParent ? _rootElement : _rootElement.FindElementByIndexInParent(index);
        }

        /// <summary>
        /// 用 children count查找特定的一个JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement FindElementByChildrenCount(int count)
        {
            if (count < 0) return null;
            return count == _rootElement.ChildrenCount ? _rootElement : _rootElement.FindElementByChildrenCount(count);
        }

        /// <summary>
        /// 用Xpath寻找指定的JabElement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXPath"></param>
        /// <returns></returns>
        public JabElement FindElementByXPath(string strXPath)
        {
            return _rootElement.FindElementByXPath(strXPath);
        }


        //==================================================


        /// <summary>
        /// 用Name寻找一组JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByName(string strName, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (_rootElement.Name == strName) result.Add(_rootElement);
            }
            else
            {
                if (new Regex(@strName).IsMatch(_rootElement.Name)) result.Add(_rootElement);
            }

            foreach (JabElement ele in _rootElement.FindElementsByName(strName, regexMatch))
            {
                result.Add(ele);
            }

            return result.ToArray();
        }

        /// <summary>
        /// 用Role寻找一组JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <param name="regexMatch"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByRole(string strRole, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (_rootElement.Role == strRole) result.Add(_rootElement);
            }
            else
            {
                if (new Regex(@strRole).IsMatch(_rootElement.Role)) result.Add(_rootElement);
            }

            foreach (JabElement ele in _rootElement.FindElementsByRole(strRole))
            {
                result.Add(ele);
            }

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
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (_rootElement.Description == strDescription) result.Add(_rootElement);
            }
            else
            {
                if (new Regex(@strDescription).IsMatch(_rootElement.Description)) result.Add(_rootElement);
            }

            foreach (JabElement ele in _rootElement.FindElementsByDescription(strDescription))
            {
                result.Add(ele);
            }

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
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (_rootElement.States == strState) result.Add(_rootElement);
            }
            else
            {
                if (new Regex(@strState).IsMatch(_rootElement.States)) result.Add(_rootElement);
            }

            foreach (JabElement ele in _rootElement.FindElementsByState(strState))
            {
                result.Add(ele);
            }

            return result.ToArray();
        }

        /// <summary>
        /// 用Object Depth查找一组JabElement
        /// </summary>
        /// <param name="objectDepth"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByObjectDepth(int objectDepth)
        {
            List<JabElement> result = new List<JabElement>();
            if (objectDepth < 0) return result.ToArray();

            if (_rootElement.ObjectDepth == objectDepth) result.Add(_rootElement);
            foreach (JabElement ele in _rootElement.FindElementsByObjectDepth(objectDepth))
            {
                result.Add(ele);
            }

            return result.ToArray();
        }

        /// <summary>
        /// 用Index In Parent 查找特定一组JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByIndexInParent(int index)
        {
            List<JabElement> result = new List<JabElement>();
            if (index < 0) return result.ToArray();

            if (_rootElement.IndexInParent == index) result.Add(_rootElement);
            foreach (JabElement ele in _rootElement.FindElementsByIndexInParent(index))
            {
                result.Add(ele);
            }

            return result.ToArray();
        }

        /// <summary>
        /// 用Children count 查找特定一组JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByChildrenCount(int count)
        {
            List<JabElement> result = new List<JabElement>();
            if (count < 0) return result.ToArray();

            if (_rootElement.ChildrenCount == count) result.Add(_rootElement);
            foreach (JabElement ele in _rootElement.FindElementsByChildrenCount(count))
            {
                result.Add(ele);
            }

            return result.ToArray();
        }

        /// <summary>
        /// 用XPath 查找特定一组JabElement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXPath"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByXPath(string strXPath)
        {
            return _rootElement.FindElementsByXPath(strXPath);
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
        /// 返回当前获得焦点的JabElement
        /// </summary>
        /// <returns></returns>
        public JabElement GetFocusedElement()
        {
            JavaObjectHandle ac;
            int vmid;
            bool result;
            result = _objBridge.Functions.GetAccessibleContextWithFocus(_hwnd, out vmid, out ac);
            if (!result || ac.IsNull) return null;
            return new JabElement(new AccessibleContextNode(_objBridge, ac), _hwnd);
        }

        /// <summary>
        /// 按键盘快捷键
        /// </summary>
        public void PerformKey(SKey shortcutkey)
        {
            Win32Api.SetTopWindow(_hwnd); //先让窗口获得焦点，摆到顶层窗口
            if (_shortcutDic.ContainsKey(shortcutkey)) SendKeys.SendWait(_shortcutDic[shortcutkey]);
        }

        /// <summary>
        /// 最小化java窗口
        /// </summary>
        public void Minimize_Window()
        {
            Win32Api._ShowWindow(_hwnd, 6); //SW_MINIMIZE
        }

        /// <summary>
        /// 最大化java窗口
        /// </summary>
        public void Maximize_Window()
        {
            Win32Api._ShowWindow(_hwnd, 3); //SW_MAXIMIZE
        }

        /// <summary>
        /// 设置窗口的长和宽
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void Set_Window_Size(int width, int height)
        {
            if (_hwnd == IntPtr.Zero) return;
            Win32Api.SetWindowSize(_hwnd, width, height);
        }

        /// <summary>
        /// 设置窗口位置 X Y轴坐标
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void Set_Window_Position(int x, int y)
        {
            if (_hwnd == IntPtr.Zero) return;
            Win32Api.SetWindowPosition(_hwnd, x, y);
        }

        public void ActivateWindow()
        {
            if (_hwnd == IntPtr.Zero) return;
            Win32Api.SetTopWindow(_hwnd);
        }

        /// <summary>
        /// 保存当前窗口的截图  jpg  jpeg  png
        /// </summary>
        /// <param name="savePath"></param>
        public void Save_ScreenShot(string savePath)
        {
            if (_hwnd == IntPtr.Zero)
            {
                throw new InvalidOperationException("Window Handle is Null!");
            }

            Bitmap bmp = Win32Api.GetScreenshot(_hwnd);
            if (savePath.ToLower().EndsWith(".png"))
            {
                bmp.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
                return;
            }

            if (savePath.ToLower().EndsWith(".jpg"))
            {
                bmp.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                return;
            }

            if (savePath.ToLower().EndsWith(".jpeg"))
            {
                bmp.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                return;
            }

            throw new InvalidOperationException("Unsupported Image File Type.\nTry .png .jpg or .jpeg");
        }


        /// <summary>
        /// 获取JavaWindow 封装这个函数 供另一个线程调用
        /// </summary>
        /// <param name="bridge"></param>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        private AccessibleWindow GetJavaWindow(AccessBridge bridge, IntPtr hwnd)
        {
            MySleep(300);
            //Debug.WriteLine(bridge.Functions.IsJavaWindow(hwnd));
            return bridge.CreateAccessibleWindow(hwnd);
        }

        /// <summary>
        /// 另一种方式休眠
        /// </summary>
        /// <param name="milliseconds"></param>
        private void MySleep(int milliseconds)
        {
            Task.Delay(milliseconds).Wait();
        }

        /// <summary>
        /// 初始化快捷键字典
        /// </summary>
        private void Init_ShortcutDic()
        {
            _shortcutDic = new Dictionary<SKey, string>()
            {
                { SKey.Copy, "^c" },
                { SKey.Cut, "^x" },
                { SKey.Paste, "^v" },
                { SKey.Select_All, "^a" },
                { SKey.ESC, "{ESC}" },

                { SKey.Oracle_Clear_Field, "{F5}" },
                { SKey.Oracle_Clear_Block, "{F7}" },
                { SKey.Oracle_Clear_Form, "{F8}" },
                { SKey.Oracle_Clear_Record, "{F6}" },

                { SKey.Oracle_Block_Menu, "^b" },
                { SKey.Oracle_Commit, "^s" },
                { SKey.Oracle_Count_Query, "{F12}" },
                { SKey.Oracle_Delete_Record, "^{UP}" },
                { SKey.Oracle_Display_Error, "+^e" },

                { SKey.Oracle_Duplicate_Field, "+{F5}" },
                { SKey.Oracle_Duplicate_Record, "+{F6}" },
                { SKey.Oracle_Edit, "^e" },
                { SKey.Oracle_Enter_Query, "{F11}" },
                { SKey.Oracle_Excute_Query, "^{F11}" },
                { SKey.Oracle_Exit, "{F4}" },
                { SKey.Oracle_Help, "^h" },
                { SKey.Oracle_Insert_Record, "^{DOWN}" },
                { SKey.Oracle_List_Of_Values, "^l" },
                { SKey.Oracle_List_Tab_Pages, "{F2}" },
                { SKey.Oracle_Update_Record, "^u" },

                { SKey.Oracle_Next_Block, "+{PGDN}" },
                { SKey.Oracle_Next_Field, "{TAB}" },
                { SKey.Oracle_Next_Primary_Key, "+{F7}" },
                { SKey.Oracle_Next_Record, "{DOWN}" },
                { SKey.Oracle_Next_Set_Of_Record, "+{F8}" },

                { SKey.Oracle_Previous_Block, "+{PGUP}" },
                { SKey.Oracle_Previous_Field, "+{TAB}" },
                { SKey.Oracle_Previous_Record, "+{UP}" },

                { SKey.Oracle_Print, "^p" },
                { SKey.Oracle_Scroll_Down, "{PGDN}" },
                { SKey.Oracle_Scroll_Up, "{PGUP}" },
                { SKey.Oracle_Up, "{UP}" },
                { SKey.Oracle_Down, "{DOWN}" },

                { SKey.Oracle_Show_Keys, "^k" },
            };
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
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
        private AccessBridge objBridge;
        private IntPtr hwnd;
        private AccessibleWindow javaWindow;
        private JabElement rootElement;
        private STAThread messageThread;
        private Win32Api api;

        private readonly Dictionary<SKey, string> _shortcut_dic;

        public IntPtr w_hwnd
        {
            get { return hwnd; }
        }

        public int Jvm_ID
        {
            get { return javaWindow.JvmId; }
        }

        public int PID
        {
            get { return api.GetPidFromHwnd(hwnd); }
        }



        /// <summary>
        /// 构造函数
        /// </summary>
        public JabDriver()
        {
            //初始化Key字典
            _shortcut_dic = new Dictionary<SKey, string>()
            {
                {SKey.Copy,"^c" },
                {SKey.Cut,"^x" },
                {SKey.Paste,"^v" },
                {SKey.Select_All,"^a"},
                {SKey.ESC,"{ESC}"},

                {SKey.Oracle_Clear_Field ,"{F5}"},
                {SKey.Oracle_Clear_Block,"{F7}"},
                {SKey.Oracle_Clear_Form,"{F8}"},
                {SKey.Oracle_Clear_Record,"{F6}"},

                {SKey.Oracle_Block_Menu,"^b"},
                {SKey.Oracle_Commit,"^s"},
                {SKey.Oracle_Count_Query,"{F12}"},
                {SKey.Oracle_Delete_Record,"^{UP}"},
                {SKey.Oracle_Display_Error,"+^e"},

                {SKey.Oracle_Duplicate_Field,"+{F5}"},
                {SKey.Oracle_Duplicate_Record,"+{F6}"},
                {SKey.Oracle_Edit,"^e"},
                {SKey.Oracle_Enter_Query,"{F11}"},
                {SKey.Oracle_Excute_Query,"^{F11}"},
                {SKey.Oracle_Exit,"{F4}"},
                {SKey.Oracle_Help,"^h"},
                {SKey.Oracle_Insert_Record,"^{DOWN}"},
                {SKey.Oracle_List_Of_Values,"^l"},
                {SKey.Oracle_List_Tab_Pages,"{F2}"},
                {SKey.Oracle_Update_Record,"^u"},

                {SKey.Oracle_Next_Block,"+{PGDN}"},
                {SKey.Oracle_Next_Field,"{TAB}"},
                {SKey.Oracle_Next_Primary_Key,"+{F7}"},
                {SKey.Oracle_Next_Record,"{DOWN}"},
                {SKey.Oracle_Next_Set_Of_Record,"+{F8}"},

                {SKey.Oracle_Previous_Block,"+{PGUP}"},
                {SKey.Oracle_Previous_Field,"+{TAB}"},
                {SKey.Oracle_Previous_Record,"+{UP}"},

                {SKey.Oracle_Print,"^p"},
                {SKey.Oracle_Scroll_Down,"{PGDN}"},
                {SKey.Oracle_Scroll_Up,"{PGUP}"},
                {SKey.Oracle_Up,"{UP}"},
                {SKey.Oracle_Down,"{DOWN}"},

                {SKey.Oracle_Show_Keys,"^k"},
            };

            javaWindow = null;
            api = new Win32Api();
            hwnd = IntPtr.Zero;
            messageThread = new STAThread();
            objBridge = new AccessBridge();
            objBridge.Initilized += (sender1, args) => { Debug.WriteLine("Initilize Bridge Successfully"); };   //Bridge初始化成功就打印这句话
            messageThread.Invoke(new Action(() => { objBridge.Initialize(); }), new object[] { });
        }

        /// <summary>
        /// 析构函数
        /// </summary>
        ~JabDriver()
        {
            try { messageThread.Dispose(); }
            catch (Exception e) { Debug.WriteLine(e.Message); }
        }

        /// <summary>
        /// 外部调用 初始化JabDriver对象
        /// </summary>
        /// <param name="title"></param>
        /// <param name="timeout"></param>
        public void Init_JabDriver(string title, int timeoutSecond = 10)
        {
            long startTick = DateTime.Now.Ticks;
            while (true)
            {
                foreach (KeyValuePair<IntPtr, string> window in api.GetAllOpenWindows())
                {
                    if (window.Value == title)
                    {
                        hwnd = window.Key;
                        break;
                    }
                }
                //hwnd = api.FindWindowByCaption(title);
                Func<AccessBridge, IntPtr, AccessibleWindow> f1 = GetJavaWindow;
                javaWindow = (AccessibleWindow)messageThread.Invoke(f1, new object[] { objBridge, hwnd });
                if (javaWindow != null) break;

                long elapsedTicks = DateTime.Now.Ticks - startTick;
                if (new TimeSpan(elapsedTicks).TotalSeconds > timeoutSecond)
                {
                    messageThread.Dispose();
                    throw new InvalidOperationException(String.Format("Can't find java window by title '{0}' in '{1}' seconds", title, timeoutSecond));
                }
            }

            api.SetTopWindow(hwnd);   //java窗口放到最前面
            rootElement = new JabElement(javaWindow, hwnd);  //根节点赋值
        }

        /// <summary>
        /// 寻找指定Name的JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public JabElement FindElementByName(string strName, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                if (strName == rootElement.name)
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByName(strName);
                }
            }
            else
            {
                if (new Regex(@strName).IsMatch(rootElement.name))
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByName(strName, regexMatch);
                }
            }

        }

        /// <summary>
        /// 寻找指定Role的JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <returns></returns>
        public JabElement FindElementByRole(string strRole, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                if (strRole == rootElement.role)
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByRole(strRole);
                }
            }
            else
            {
                if (new Regex(@strRole).IsMatch(rootElement.role))
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByRole(strRole, regexMatch);
                }
            }

        }

        /// <summary>
        /// 寻找指定Description的JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <returns></returns>
        public JabElement FindElementByDescription(string strDescription, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                if (strDescription == rootElement.description)
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByDescription(strDescription);
                }
            }
            else
            {
                if (new Regex(@strDescription).IsMatch(rootElement.description))
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByDescription(strDescription, regexMatch);
                }
            }

        }

        /// <summary>
        /// 用State查找特定一个JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <returns></returns>
        public JabElement FindElementByState(string strState, bool regexMatch = false)
        {
            if (!regexMatch)
            {
                if (strState == rootElement.states)
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByState(strState);
                }
            }
            else
            {
                if (new Regex(@strState).IsMatch(rootElement.states))
                {
                    return rootElement;
                }
                else
                {
                    return rootElement.FindElementByState(strState, regexMatch);
                }
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
            if (objectDepth == rootElement.object_depth)
            {
                return rootElement;
            }
            else
            {
                return rootElement.FindElementByObjectDepth(objectDepth);
            }
        }

        /// <summary>
        /// 用 index in parent 查找特定的一个JabElement
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public JabElement FindElementByIndexInParent(int index)
        {
            if (index < 0) return null;
            if (index == rootElement.index_in_parent)
            {
                return rootElement;
            }
            else
            {
                return rootElement.FindElementByIndexInParent(index);
            }
        }

        /// <summary>
        /// 用 children count查找特定的一个JabElement
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public JabElement FindElementByChildrenCount(int count)
        {
            if (count < 0) return null;
            if (count == rootElement.children_count)
            {
                return rootElement;
            }
            else
            {
                return rootElement.FindElementByChildrenCount(count);
            }
        }

        /// <summary>
        /// 用Xpath寻找指定的JabElement
        /// 应尽量少用Xpath ，每次调用都要遍历所有节点，生成xml，开销比较大
        /// </summary>
        /// <param name="strXPath"></param>
        /// <returns></returns>
        public JabElement FindElementByXPath(string strXPath)
        {
            return rootElement.FindElementByXPath(strXPath);
        }



        //==================================================


        /// <summary>
        /// 用Name寻找一组JabElement
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByName(string strName, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (rootElement.name == strName) result.Add(rootElement);
            }
            else
            {
                if (new Regex(@strName).IsMatch(rootElement.name)) result.Add(rootElement);
            }

            foreach (JabElement ele in rootElement.FindElementsByName(strName, regexMatch))
            {
                result.Add(ele);
            }
            return result.ToArray();
        }

        /// <summary>
        /// 用Role寻找一组JabElement
        /// </summary>
        /// <param name="strRole"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByRole(string strRole, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (rootElement.role == strRole) result.Add(rootElement);
            }
            else
            {
                if (new Regex(@strRole).IsMatch(rootElement.role)) result.Add(rootElement);
            }
            foreach (JabElement ele in rootElement.FindElementsByRole(strRole))
            {
                result.Add(ele);
            }
            return result.ToArray();
        }

        /// <summary>
        /// 用Description查找一组JabElement
        /// </summary>
        /// <param name="strDescription"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByDescription(string strDescription, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (rootElement.description == strDescription) result.Add(rootElement);
            }
            else
            {
                if (new Regex(@strDescription).IsMatch(rootElement.description)) result.Add(rootElement);
            }
            foreach (JabElement ele in rootElement.FindElementsByDescription(strDescription))
            {
                result.Add(ele);
            }
            return result.ToArray();
        }

        /// <summary>
        /// 用State查找一组JabElement
        /// </summary>
        /// <param name="strState"></param>
        /// <returns></returns>
        public JabElement[] FindElementsByState(string strState, bool regexMatch = false)
        {
            List<JabElement> result = new List<JabElement>();
            if (!regexMatch)
            {
                if (rootElement.states == strState) result.Add(rootElement);
            }
            else
            {
                if (new Regex(@strState).IsMatch(rootElement.states)) result.Add(rootElement);
            }
            foreach (JabElement ele in rootElement.FindElementsByState(strState))
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

            if (rootElement.object_depth == objectDepth) result.Add(rootElement);
            foreach (JabElement ele in rootElement.FindElementsByObjectDepth(objectDepth))
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

            if (rootElement.index_in_parent == index) result.Add(rootElement);
            foreach (JabElement ele in rootElement.FindElementsByIndexInParent(index))
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

            if (rootElement.children_count == count) result.Add(rootElement);
            foreach (JabElement ele in rootElement.FindElementsByChildrenCount(count))
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
            return rootElement.FindElementsByXPath(strXPath);
        }

        /// <summary>
        /// 设置一个超时的秒数，一直寻找一个特定的JabElement
        /// </summary>
        /// <param name="by"></param>
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
        /// 返回当前获得焦点的JabElement
        /// </summary>
        /// <returns></returns>
        public JabElement GetFocusedElement()
        {
            JavaObjectHandle ac;
            int vmid;
            bool result;
            result = objBridge.Functions.GetAccessibleContextWithFocus(hwnd, out vmid, out ac);
            if (!result || ac.IsNull) return null;
            return new JabElement(new AccessibleContextNode(objBridge, ac), hwnd);
        }

        /// <summary>
        /// 按键盘快捷键
        /// </summary>
        public void PerformKey(SKey shortcutkey)
        {
            api.SetTopWindow(hwnd); //先让窗口获得焦点，摆到顶层窗口
            if (_shortcut_dic.ContainsKey(shortcutkey)) SendKeys.SendWait(_shortcut_dic[shortcutkey]);
        }

        /// <summary>
        /// 最小化java窗口
        /// </summary>
        public void Minimize_Window()
        {
            api._ShowWindow(hwnd, 6); //SW_MINIMIZE
        }

        /// <summary>
        /// 最大化java窗口
        /// </summary>
        public void Maximize_Window()
        {
            api._ShowWindow(hwnd, 3);  //SW_MAXIMIZE
        }

        /// <summary>
        /// 设置窗口的长和宽
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void Set_Window_Size(int width, int height)
        {
            if (hwnd == IntPtr.Zero) return;
            api.SetWindowSize(hwnd, width, height);
        }

        /// <summary>
        /// 设置窗口位置 X Y轴坐标
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void Set_Window_Position(int x, int y)
        {
            if (hwnd == IntPtr.Zero) return;
            api.SetWindowPosition(hwnd, x, y);
        }

        public void ActivateWindow()
        {
            if (hwnd == IntPtr.Zero) return;
            api.SetTopWindow(hwnd);
        }

        /// <summary>
        /// 保存当前窗口的截图  jpg  jpeg  png
        /// </summary>
        /// <param name="savePath"></param>
        public void Save_ScreenShot(string savePath)
        {
            if (hwnd == IntPtr.Zero)
            {
                throw new InvalidOperationException("Window Handle is Null!");
            }
            Bitmap bmp = api.GetScreenshot(hwnd);
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
            throw new InvalidOperationException("Unsupport Image File Type.\nTry .png .jpg or .jpeg");
        }



        /// <summary>
        /// 获取JavaWindow 封装这个函数 供另一个线程调用
        /// </summary>
        /// <param name="bridge"></param>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        private AccessibleWindow GetJavaWindow(AccessBridge bridge, IntPtr hwnd)
        {
            Thread.Sleep(300);
            //Debug.WriteLine(bridge.Functions.IsJavaWindow(hwnd));
            return bridge.CreateAccessibleWindow(hwnd);
        }



    }
}

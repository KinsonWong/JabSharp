# JabSharp
##### 用C#写的一个对Java Access Bridge的简单包装

主要的类：JabDriver，JabElement

当需要对Java窗体元素进行定位，调用的方法与Selenium十分类似,例如FindElementByName，FindElementByRole

另外引入Xpath查找功能，实现方法是将Java窗体所有元素的信息写入到XML中，再调用Xpath进行定位。

相关工具
[access-bridge-explorer](https://github.com/google/access-bridge-explorer)
用于检索Java窗体元素，作用类似于Spy


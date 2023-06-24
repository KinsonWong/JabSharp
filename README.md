# JabSharp
##### 用C#写的一个对Java Access Bridge的简单包装
主要的类：JabDriver，JabElement <br>
当需要对Java窗体元素进行定位，调用的方法与Selenium十分类似，例如FindElementByName，FindElementByRole <br>
另外引入Xpath查找功能，实现方法是将Java窗体所有元素的信息写入到XML中，再调用Xpath进行定位。 <br>

---

### 启用Java Access Bridge
进入Java根目录的bin文件夹中<br>
cmd输入以下命令，需要重启后生效<br>
`jabswitch.exe -enable` 

---

### 相关工具
[access-bridge-explorer](https://github.com/google/access-bridge-explorer) <br>
用于检索Java窗体元素，作用类似于Spy

---

### 更新
2023/06/24 <br>
使用Xunit框架，编写几个简单的测试用例，对项目进行简单的测试 <br>

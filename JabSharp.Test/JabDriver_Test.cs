using System;
using BasicJab;
using Xunit;
using Xunit.Abstractions;

namespace JabSharp.Test
{
    public class JabDriver_Test:IDisposable
    {
        private readonly ITestOutputHelper _outputHelper;
        private JabDriver _driver;

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="testOutputHelper"></param>
        public JabDriver_Test(ITestOutputHelper testOutputHelper)
        {
            _outputHelper = testOutputHelper;
            _driver = new JabDriver();
            _driver.Init_JabDriver("Java Control Panel", 3);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            _driver = null;
        }

        [Theory]
        [InlineData("XXXX", 1)]
        //找不到 java窗口 抛出异常
        public void Init_Driver_Test(string title, int timeout)
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                try
                {
                    var driver = new JabDriver();
                    driver.Init_JabDriver(title, timeout);
                }
                catch (Exception ex)
                {
                    _outputHelper.WriteLine(ex.Message);
                    throw;
                }
            });
        }

        [Theory]
        [InlineData("testName", false)]
        [InlineData("General", false)]
        [InlineData("(?i)general", true)]
        public void FindElementByName_Test(string name, bool isRegex)
        {
            var ele = _driver.FindElementByName(name, isRegex);
            if (!isRegex)
            {
                if (ele != null) Assert.Equal(ele.Name, name);
            }
            else
            {
                if (ele != null) Assert.Matches(name, ele.Name);
            }
        }

        [Theory]
        [InlineData("//page_tab[@name='General']", "General")]
        public void FindElementByXPath_Test(string strXPath, string checkName)
        {
            var ele = _driver.FindElementByXPath(strXPath);
            if (ele != null) Assert.Equal(ele.Name, checkName);
        }

        [Theory]
        [InlineData(By.NAME, "General", false)]
        [InlineData(By.NAME, "(?i)general", true)]
        public void WaitUntil_Test(By byWhat, object value, bool isRegex)
        {
            var ele = _driver.WaitUntilElementExists(byWhat, value, isRegex);
            if (!isRegex)
            {
                if (ele != null) Assert.Equal(value.ToString(), ele.Name);
            }
            else
            {
                if (ele != null) Assert.Matches(value.ToString(), ele.Name);
            }
        }

        [Theory]
        [InlineData("(?i)general", true, 2)]
        [InlineData("xxx", false, 0)]
        public void FindElementsByName_Test(string strName, bool isRegex, int expectNum)
        {
            var ele = _driver.FindElementsByName(strName, isRegex);
            Assert.Equal(ele.Length,expectNum);
        }
    }
}
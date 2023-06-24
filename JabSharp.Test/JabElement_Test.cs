using System;
using BasicJab;
using Xunit;
using Xunit.Abstractions;

namespace JabSharp.Test
{
    public class JabElement_Test : IDisposable
    {
        private readonly ITestOutputHelper _outputHelper;
        private JabDriver _driver;

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="testOutputHelper"></param>
        public JabElement_Test(ITestOutputHelper testOutputHelper)
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

        [Fact]
        public void Click_Test()
        {
            var ele = _driver.FindElementByXPath("//page_tab[@name='Advanced']");
            if (ele == null) return;

            _driver.ActivateWindow();
            ele.Click(true);
        }
    }
}
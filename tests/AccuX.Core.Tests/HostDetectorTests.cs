using System;
using AccuX.Core.Operations;
using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// 宿主识别逻辑测试。DetectKind 是纯字符串判断，可脱离 COM 测试。
    /// </summary>
    public class HostDetectorTests
    {
        [Theory]
        [InlineData("Microsoft Excel", @"C:\Program Files\Microsoft Office\root\Office16", HostKind.Excel)]
        [InlineData("WPS 表格", @"C:\Users\me\AppData\Local\Kingsoft\WPS Office", HostKind.Wps)]
        [InlineData("ET", @"C:\Program Files\Kingsoft\WPS Office", HostKind.Wps)]
        [InlineData("", "", HostKind.Unknown)]
        public void DetectKind_ClassifiesByProbe(string name, string path, HostKind expected)
        {
            Assert.Equal(expected, HostDetector.DetectKind(name, path));
        }

        [Fact]
        public void Detect_Null_ReturnsUnknown()
        {
            var context = HostDetector.Detect(null);

            Assert.Equal(HostKind.Unknown, context.HostKind);
            Assert.Equal(IntPtr.Zero, context.MainWindowHandle);
        }
    }
}

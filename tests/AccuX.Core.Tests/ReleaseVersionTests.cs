using AccuX.Core.Updates;
using Xunit;

namespace AccuX.Core.Tests
{
    public sealed class ReleaseVersionTests
    {
        [Theory]
        [InlineData("v1.4", 1, 4)]
        [InlineData("1.4", 1, 4)]
        [InlineData("v0.0", 0, 0)]
        public void TryParse_AcceptsStableTwoPartTags(string text, int major, int minor)
        {
            Assert.True(ReleaseVersion.TryParse(text, out var version));
            Assert.Equal(major, version.Major);
            Assert.Equal(minor, version.Minor);
            Assert.Equal(major + "." + minor, version.Text);
        }

        [Theory]
        [InlineData("v1.4.1")]
        [InlineData("v1.4-beta")]
        [InlineData("v01.4")]
        [InlineData("release-1.4")]
        [InlineData("")]
        public void TryParse_RejectsUnsupportedTags(string text)
        {
            Assert.False(ReleaseVersion.TryParse(text, out _));
        }

        [Fact]
        public void CompareTo_ComparesNumbersInsteadOfText()
        {
            Assert.True(ReleaseVersion.TryParse("v1.10", out var ten));
            Assert.True(ReleaseVersion.TryParse("v1.4", out var four));

            Assert.True(ten.CompareTo(four) > 0);
        }
    }
}

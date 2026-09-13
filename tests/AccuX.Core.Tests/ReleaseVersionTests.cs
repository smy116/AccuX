using AccuX.Core.Updates;
using Xunit;

namespace AccuX.Core.Tests
{
    public sealed class ReleaseVersionTests
    {
        [Theory]
        [InlineData("v1.4", 1, 4, 0)]
        [InlineData("1.4", 1, 4, 0)]
        [InlineData("v0.0", 0, 0, 0)]
        [InlineData("v1.6.1", 1, 6, 1)]
        [InlineData("1.6.1", 1, 6, 1)]
        [InlineData("V1.6.10", 1, 6, 10)]
        public void TryParse_AcceptsStableTwoOrThreePartTags(string text, int major, int minor, int patch)
        {
            Assert.True(ReleaseVersion.TryParse(text, out var version));
            Assert.Equal(major, version.Major);
            Assert.Equal(minor, version.Minor);
            Assert.Equal(patch, version.Patch);
        }

        [Theory]
        [InlineData("v1.4", "1.4")]
        [InlineData("v1.6.1", "1.6.1")]
        public void Text_PreservesWhetherPatchWasSpecified(string text, string expected)
        {
            Assert.True(ReleaseVersion.TryParse(text, out var version));

            Assert.Equal(expected, version.Text);
        }

        [Theory]
        [InlineData("v1.4-beta")]
        [InlineData("v1.4.1-beta")]
        [InlineData("v01.4")]
        [InlineData("v1.6.01")]
        [InlineData("v1.4.1.2")]
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

        [Fact]
        public void CompareTo_TreatsMissingPatchAsZero()
        {
            Assert.True(ReleaseVersion.TryParse("v1.6", out var minorOnly));
            Assert.True(ReleaseVersion.TryParse("v1.6.0", out var explicitZero));
            Assert.True(ReleaseVersion.TryParse("v1.6.1", out var patchOne));

            // 1.6 与 1.6.0 视为同一版本，这是升级检测判定“已是最新”的基础。
            Assert.Equal(0, minorOnly.CompareTo(explicitZero));
            Assert.Equal(0, explicitZero.CompareTo(minorOnly));

            Assert.True(patchOne.CompareTo(minorOnly) > 0);
            Assert.True(minorOnly.CompareTo(patchOne) < 0);
        }

        [Theory]
        [InlineData("1.6.1-ci.37.d202798", 1, 6, 1, "ci.37.d202798")]
        [InlineData("1.6-ci.37.d202798", 1, 6, 0, "ci.37.d202798")]
        [InlineData("1.6.1+abc1234", 1, 6, 1, null)]
        [InlineData("1.6.1.0", 1, 6, 1, null)]
        [InlineData("v0.0.0-dev", 0, 0, 0, "dev")]
        public void TryParseLoose_ReadsCurrentVersionShapes(
            string text, int major, int minor, int patch, string prerelease)
        {
            Assert.True(ReleaseVersion.TryParseLoose(text, out var version));
            Assert.Equal(major, version.Major);
            Assert.Equal(minor, version.Minor);
            Assert.Equal(patch, version.Patch);
            Assert.Equal(prerelease, version.Prerelease);
            Assert.Equal(prerelease != null, version.IsPrerelease);
        }

        [Fact]
        public void CompareTo_RanksPrereleaseBelowSameStableVersion()
        {
            Assert.True(ReleaseVersion.TryParseLoose("1.6.1-ci.37.d202798", out var preview));
            Assert.True(ReleaseVersion.TryParse("v1.6.1", out var stable));
            Assert.True(ReleaseVersion.TryParse("v1.6", out var older));

            // 预发布小于同号正式版，才能让测试版用户收到正式版的升级提示；
            // 同时大于上一个正式版，避免被回退。
            Assert.True(stable.CompareTo(preview) > 0);
            Assert.True(preview.CompareTo(stable) < 0);
            Assert.True(preview.CompareTo(older) > 0);
        }
    }
}

using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AccuX.Core.Updates
{
    /// <summary>
    /// AccuX Release 使用的版本号，支持两段（例如 1.6）或三段（例如 1.6.1），
    /// 以及可选的预发布后缀（例如 CI 构建的 1.6.1-ci.37.d202798）。
    ///
    /// 比较遵循 SemVer 的排序：缺失的修订号按 0 处理（1.6 == 1.6.0），
    /// 同一数字版本下预发布小于正式版（1.6.1-ci.37 &lt; 1.6.1）。后一条是
    /// 预发布用户能自然收敛到正式版的基础。正式 Release 的 tag 不含预发布后缀。
    /// </summary>
    public sealed class ReleaseVersion : IComparable<ReleaseVersion>
    {
        // 正式 tag：只允许 主.次 或 主.次.修订。
        private static readonly Regex TagPattern = new Regex(
            "^v?(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)(?:\\.(0|[1-9][0-9]*))?$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        // 当前运行版本：允许预发布后缀与构建元数据（+...，由调用方先行剥离）。
        private static readonly Regex LoosePattern = new Regex(
            "^v?(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)(?:\\.(0|[1-9][0-9]*))?(?:-([0-9A-Za-z.\\-]+))?$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public ReleaseVersion(int major, int minor)
            : this(major, minor, 0, null, false)
        {
        }

        public ReleaseVersion(int major, int minor, int patch)
            : this(major, minor, patch, null, patch != 0)
        {
        }

        public ReleaseVersion(int major, int minor, int patch, string prerelease)
            : this(major, minor, patch, prerelease, patch != 0)
        {
        }

        private ReleaseVersion(int major, int minor, int patch, string prerelease, bool hasPatch)
        {
            if (major < 0 || minor < 0 || patch < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            Major = major;
            Minor = minor;
            Patch = patch;
            Prerelease = string.IsNullOrWhiteSpace(prerelease) ? null : prerelease.Trim();
            HasPatch = hasPatch;
        }

        public int Major { get; }

        public int Minor { get; }

        public int Patch { get; }

        /// <summary>预发布标识（不含前导 '-'）；正式版为 null。</summary>
        public string Prerelease { get; }

        public bool IsPrerelease
        {
            get { return Prerelease != null; }
        }

        /// <summary>
        /// 解析来源是否显式给出了修订号。只影响 <see cref="Text"/> 的显示形态，
        /// 不参与比较（1.6 与 1.6.0 视为同一版本）。
        /// </summary>
        public bool HasPatch { get; }

        public string Text
        {
            get
            {
                var text = HasPatch
                    ? Major.ToString(CultureInfo.InvariantCulture)
                        + "." + Minor.ToString(CultureInfo.InvariantCulture)
                        + "." + Patch.ToString(CultureInfo.InvariantCulture)
                    : Major.ToString(CultureInfo.InvariantCulture)
                        + "." + Minor.ToString(CultureInfo.InvariantCulture);
                return IsPrerelease ? text + "-" + Prerelease : text;
            }
        }

        public int CompareTo(ReleaseVersion other)
        {
            if (other == null)
            {
                return 1;
            }

            var numeric = Major.CompareTo(other.Major);
            if (numeric != 0)
            {
                return numeric;
            }

            numeric = Minor.CompareTo(other.Minor);
            if (numeric != 0)
            {
                return numeric;
            }

            numeric = Patch.CompareTo(other.Patch);
            if (numeric != 0)
            {
                return numeric;
            }

            // 数字部分相同：预发布小于正式版。
            if (IsPrerelease && !other.IsPrerelease)
            {
                return -1;
            }

            if (!IsPrerelease && other.IsPrerelease)
            {
                return 1;
            }

            return IsPrerelease
                ? string.CompareOrdinal(Prerelease, other.Prerelease)
                : 0;
        }

        /// <summary>
        /// 解析正式 Release 的 tag（严格：两段或三段，不含预发布）。
        /// </summary>
        public static bool TryParse(string value, out ReleaseVersion version)
        {
            version = null;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var match = TagPattern.Match(value.Trim());
            if (!match.Success || !TryReadNumeric(match, out var major, out var minor, out var patch, out var hasPatch))
            {
                return false;
            }

            version = new ReleaseVersion(major, minor, patch, null, hasPatch);
            return true;
        }

        /// <summary>
        /// 解析客户端当前版本（宽松：允许预发布后缀与四段数字，用于升级检测）。
        /// 构建元数据 "+..." 会被忽略。
        /// </summary>
        public static bool TryParseLoose(string value, out ReleaseVersion version)
        {
            version = null;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var text = value.Trim();
            var plus = text.IndexOf('+');
            if (plus >= 0)
            {
                text = text.Substring(0, plus);
            }

            var match = LoosePattern.Match(text);
            if (match.Success && TryReadNumeric(match, out var major, out var minor, out var patch, out var hasPatch))
            {
                version = new ReleaseVersion(major, minor, patch, match.Groups[4].Value, hasPatch);
                return true;
            }

            // 回退：容忍四段程序集版本，例如 1.6.1.0，取前三段。
            var fallback = Regex.Match(text, "^(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)(?:\\.(0|[1-9][0-9]*))?");
            if (!fallback.Success)
            {
                return false;
            }

            if (!int.TryParse(fallback.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var fallbackMajor)
                || !int.TryParse(fallback.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var fallbackMinor))
            {
                return false;
            }

            var fallbackPatch = 0;
            var fallbackHasPatch = false;
            if (fallback.Groups[3].Success)
            {
                if (!int.TryParse(fallback.Groups[3].Value, NumberStyles.None, CultureInfo.InvariantCulture, out fallbackPatch))
                {
                    return false;
                }

                fallbackHasPatch = true;
            }

            version = new ReleaseVersion(fallbackMajor, fallbackMinor, fallbackPatch, null, fallbackHasPatch);
            return true;
        }

        private static bool TryReadNumeric(
            Match match,
            out int major,
            out int minor,
            out int patch,
            out bool hasPatch)
        {
            major = 0;
            minor = 0;
            patch = 0;
            hasPatch = false;

            if (!int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out major)
                || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out minor))
            {
                return false;
            }

            if (!match.Groups[3].Success)
            {
                return true;
            }

            if (!int.TryParse(match.Groups[3].Value, NumberStyles.None, CultureInfo.InvariantCulture, out patch))
            {
                return false;
            }

            hasPatch = true;
            return true;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}

using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AccuX.Core.Updates
{
    /// <summary>
    /// AccuX 稳定 Release 使用的两段式版本号，例如 1.4。
    /// </summary>
    public sealed class ReleaseVersion : IComparable<ReleaseVersion>
    {
        private static readonly Regex TagPattern = new Regex(
            "^v?(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public ReleaseVersion(int major, int minor)
        {
            if (major < 0 || minor < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            Major = major;
            Minor = minor;
        }

        public int Major { get; }

        public int Minor { get; }

        public string Text
        {
            get { return Major.ToString(CultureInfo.InvariantCulture) + "." + Minor.ToString(CultureInfo.InvariantCulture); }
        }

        public int CompareTo(ReleaseVersion other)
        {
            if (other == null)
            {
                return 1;
            }

            var major = Major.CompareTo(other.Major);
            return major != 0 ? major : Minor.CompareTo(other.Minor);
        }

        public static bool TryParse(string value, out ReleaseVersion version)
        {
            version = null;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var match = TagPattern.Match(value.Trim());
            if (!match.Success
                || !int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var major)
                || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var minor))
            {
                return false;
            }

            version = new ReleaseVersion(major, minor);
            return true;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}

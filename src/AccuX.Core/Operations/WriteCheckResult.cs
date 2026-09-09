using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 写入前可写性检查结果（规格 §5.2 / §18 / §19）。
    /// 业务层据此判断是否继续写回，不依赖 COM Exception。
    /// </summary>
    public sealed class WriteCheckResult
    {
        private readonly List<string> _issues = new List<string>();

        public WriteCheckResult(bool canWrite)
        {
            CanWrite = canWrite;
        }

        public bool CanWrite { get; private set; }

        /// <summary>阻止写入的原因描述，用于向用户提示。</summary>
        public IReadOnlyList<string> Issues
        {
            get { return _issues; }
        }

        public static WriteCheckResult Success()
        {
            return new WriteCheckResult(true);
        }

        public static WriteCheckResult Failure(params string[] issues)
        {
            var result = new WriteCheckResult(false);
            if (issues != null)
            {
                foreach (var issue in issues)
                {
                    if (!string.IsNullOrWhiteSpace(issue))
                    {
                        result._issues.Add(issue);
                    }
                }
            }

            return result;
        }

        public WriteCheckResult AddIssue(string issue)
        {
            if (!string.IsNullOrWhiteSpace(issue))
            {
                _issues.Add(issue);
                CanWrite = false;
            }

            return this;
        }
    }
}

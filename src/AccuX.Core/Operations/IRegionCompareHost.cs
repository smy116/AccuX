using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 区域存在对比的宿主窄接口。模块只传递 RangeTarget 和 CLR 坐标，不接触 COM。
    /// </summary>
    public interface IRegionCompareHost
    {
        RangeTarget CaptureCurrentTarget();

        RangeReadResult ReadRegion(RangeTarget target);

        bool ValidateTarget(RangeTarget target, bool requireWritable, out string message);

        bool ValidatePositions(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions,
            out string message);

        long ApplyBackgroundColor(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions,
            string hexColor);

        long ClearBackgroundColor(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions);

        void ExportResults(RegionCompareExportData data);

        /// <summary>窗口关闭或插件卸载时释放本次对比会话保存的宿主身份引用。</summary>
        void ReleaseSession();
    }
}

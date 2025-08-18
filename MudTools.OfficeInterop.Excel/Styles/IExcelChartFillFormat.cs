//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Fill 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.FillFormat 或 ChartFillFormat 的安全访问和操作
/// 用于设置形状或图表元素的背景填充
/// </summary>
public interface IExcelChartFillFormat : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取填充所在的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取填充对象所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 填充属性
    /// <summary>
    /// 获取或设置填充的前景色 (RGB 颜色值)
    /// 对应 FillFormat.ForeColor 或 ChartFillFormat.ForeColor
    /// </summary>
    int ForeColor { get; }

    /// <summary>
    /// 获取或设置填充的背景色 (RGB 颜色值)
    /// 对应 FillFormat.BackColor 或 ChartFillFormat.BackColor
    /// </summary>
    int BackColor { get; }

    /// <summary>
    /// 获取或设置填充类型
    /// 对应 FillFormat.Type (使用 MsoFillType 枚举对应的 int 值)
    /// </summary>
    MsoFillType? FillType { get; }

    /// <summary>
    /// 获取或设置图案类型 (如果 FillType 为 msoFillPatterned)
    /// 对应 FillFormat.Pattern (使用 MsoPatternType 枚举对应的 int 值)
    /// </summary>
    MsoPatternType? Pattern { get; }

    // --- 渐变填充属性 (占位符) ---
    // 这些属性较为复杂，需要更细致的封装
    /// <summary>
    /// 获取或设置渐变填充的样式
    /// </summary>
    MsoGradientStyle? GradientStyle { get; }
    /// <summary>
    /// 获取或设置渐变填充的变体
    /// </summary>
    int GradientVariant { get; }
    /// <summary>
    /// 获取或设置渐变填充的颜色类型
    /// </summary>
    MsoGradientColorType? GradientColorType { get; }
    // /// <summary>
    // /// 设置指定索引的渐变停止点颜色
    // /// </summary>
    // /// <param name="index">停止点索引 (1-based)</param>
    // /// <param name="color">RGB 颜色值</param>
    // void SetGradientStopColor(int index, int color);
    // /// <summary>
    // /// 设置指定索引的渐变停止点位置
    // /// </summary>
    // /// <param name="index">停止点索引 (1-based)</param>
    // /// <param name="position">位置 (0.0 到 1.0)</param>
    // void SetGradientStopPosition(int index, float position);
    #endregion

    #region 操作方法
    /// <summary>
    /// 将填充设置为纯色
    /// 对应 FillFormat.Solid 方法
    /// </summary>
    void SetSolid();

    /// <summary>
    /// 将填充设置为无填充
    /// 对应 FillFormat.Visible = MsoTriState.msoFalse
    /// </summary>
    void SetNoFill();
    #endregion
}

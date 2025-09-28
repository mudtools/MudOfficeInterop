namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定数字在文档中的间距类型
/// </summary>
public enum WdNumberSpacing
{
    /// <summary>
    /// 默认数字间距 - 使用字体的默认设置
    /// </summary>
    wdNumberSpacingDefault,
    /// <summary>
    /// 比例数字间距 - 每个数字根据其自然宽度占用空间
    /// </summary>
    wdNumberSpacingProportional,
    /// <summary>
    /// 制表数字间距 - 所有数字占用相同的宽度，对齐显示
    /// </summary>
    wdNumberSpacingTabular
}
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据验证警告样式类型
/// </summary>
public enum XlDVAlertStyle
{
    /// <summary>
    /// 停止警告样式 - 当用户输入无效数据时显示停止图标
    /// </summary>
    xlValidAlertStop = 1,

    /// <summary>
    /// 警告样式 - 当用户输入无效数据时显示警告图标
    /// </summary>
    xlValidAlertWarning,

    /// <summary>
    /// 信息样式 - 当用户输入无效数据时显示信息图标
    /// </summary>
    xlValidAlertInformation
}
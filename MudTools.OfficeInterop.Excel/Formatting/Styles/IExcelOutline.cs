//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Outline 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Outline (Worksheet.Outline) 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelOutline : IOfficeObject<IExcelOutline>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取大纲对象的父对象 (通常是 Worksheet)
    /// 对应 Outline.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取大纲对象所在的Application对象
    /// 对应 Outline.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置自动创建大纲时是否同时创建行大纲和列大纲
    /// 对应 Outline.AutomaticStyles 属性
    /// </summary>
    bool AutomaticStyles { get; set; }

    /// <summary>
    /// 获取或设置汇总行是否显示在明细行下方
    /// 对应 Outline.SummaryRowBelow 属性
    /// </summary>
    XlSummaryRow SummaryRow { get; set; }

    /// <summary>
    /// 获取或设置汇总列是否显示在明细列右侧
    /// 对应 Outline.SummaryColumn 属性
    /// </summary>
    XlSummaryColumn SummaryColumn { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 展开指定级别的行或列大纲
    /// </summary>
    /// <param name="rowLevels">要展开的行大纲级别 (0 表示全部)</param>
    /// <param name="columnLevels">要展开的列大纲级别 (0 表示全部)</param>
    void ShowLevels(int rowLevels = 0, int columnLevels = 0);
    #endregion
}

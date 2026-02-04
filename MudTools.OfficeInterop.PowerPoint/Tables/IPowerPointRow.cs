//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中的行对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointRow : IOfficeObject<IPowerPointRow, MsPowerPoint.Row>, IDisposable
{
    /// <summary>
    /// 获取行所属的 PowerPoint 应用程序实例。
    /// </summary>
    /// <returns>PowerPoint 应用程序对象。</returns>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取行的父对象。
    /// </summary>
    /// <returns>父对象。</returns>
    object? Parent { get; }

    /// <summary>
    /// 获取行中单元格的范围。
    /// </summary>
    /// <returns>单元格范围对象。</returns>
    IPowerPointCellRange? Cells { get; }

    /// <summary>
    /// 选中当前行。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除当前行。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取或设置行的高度。
    /// </summary>
    /// <value>行的高度值（单位由 PowerPoint 应用程序决定）。</value>
    float Height { get; set; }
}
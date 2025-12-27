//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中链接对象（如链接图片、OLE对象）的链接格式设置接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.LinkFormat
/// 用于管理链接源、更新方式、断开链接等操作。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelLinkFormat : IOfficeObject<IExcelLinkFormat>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape 或 OLEObject）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置链接是否自动更新。
    /// </summary>
    bool AutoUpdate { get; set; }

    /// <summary>
    /// 获取或设置链接对象是否被锁定。
    /// 当设置为 true 时，链接对象将被锁定，防止被修改。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 立即更新链接内容（从源文件重新加载）。
    /// </summary>
    void Update();

}
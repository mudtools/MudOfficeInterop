
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中链接对象（如链接图片、OLE对象）的链接格式设置接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.LinkFormat
/// 用于管理链接源、更新方式、断开链接等操作。
/// </summary>
public interface IExcelLinkFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape 或 OLEObject）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

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
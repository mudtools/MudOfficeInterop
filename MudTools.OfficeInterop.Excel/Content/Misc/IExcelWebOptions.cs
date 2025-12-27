//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 包含将文档另存为网页或打开网页时，Microsoft Excel使用的工作簿级属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelWebOptions : IOfficeObject<IExcelWebOptions>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示在Web浏览器中查看已保存文档时是否使用级联样式表进行字体格式化。
    /// True表示使用CSS，False表示使用HTML &lt;FONT&gt;标签和CSS。默认值为True。
    /// </summary>
    bool RelyOnCSS { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将指定文档另存为网页时，是否将所有支持文件（如背景纹理和图形）组织在单独的文件夹中。
    /// True表示支持文件保存在单独文件夹，False表示支持文件与网页保存在同一文件夹。默认值为True。
    /// </summary>
    bool OrganizeInFolder { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将文档另存为网页时是否使用长文件名。
    /// True表示使用长文件名，False表示不使用长文件名，使用DOS文件名格式（8.3）。默认值为True。
    /// </summary>
    bool UseLongFileNames { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示在Web浏览器中查看已保存文档时，如果尚未安装Microsoft Office Web组件，是否下载它们。
    /// True表示下载，False表示不下载。默认值为False。
    /// </summary>
    bool DownloadComponents { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将文档另存为网页时是否不从绘图对象生成图像文件。
    /// True表示不生成图像，False表示生成图像。默认值为False。
    /// </summary>
    bool RelyOnVML { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将文档另存为网页时是否允许PNG（便携式网络图形）作为图像格式。
    /// True表示允许PNG，False表示不允许。默认值为False。
    /// </summary>
    bool AllowPNG { get; set; }

    /// <summary>
    /// 获取或设置一个MsoScreenSize常量，指定在Web浏览器中查看已保存文档时应使用的最佳最小屏幕尺寸（宽度x高度，以像素为单位）。
    /// 默认常量为msoScreenSize800x600。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoScreenSize ScreenSize { get; set; }

    /// <summary>
    /// 获取或设置网页上图形图像和表格单元格的密度（每英寸像素数）。
    /// 设置范围通常为19到480，常用设置包括72、96和120。默认设置为96。
    /// </summary>
    int PixelsPerInch { get; set; }

    /// <summary>
    /// 获取或设置中心URL（在intranet或Web上）或路径（本地或网络），授权用户在查看已保存文档时可以从此处下载Microsoft Office Web组件。
    /// 默认值是Microsoft Office的本地或网络安装路径。
    /// </summary>
    string LocationOfComponents { get; set; }

    /// <summary>
    /// 获取或设置在Web浏览器中查看已保存文档时要使用的文档编码（代码页或字符集）。
    /// 默认值是系统代码页。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding Encoding { get; set; }

    /// <summary>
    /// 获取当将文档另存为网页、使用长文件名并选择将支持文件保存在单独文件夹时，Excel使用的文件夹后缀。
    /// （即当UseLongFileNames和OrganizeInFolder属性设置为True时）。
    /// </summary>
    string FolderSuffix { get; }

    /// <summary>
    /// 将指定文档的文件夹后缀设置为所选或已安装语言支持的默认后缀。
    /// </summary>
    void UseDefaultFolderSuffix();

    /// <summary>
    /// 获取或设置一个MsoTargetBrowser常量，指示浏览器版本。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTargetBrowser TargetBrowser { get; set; }
}

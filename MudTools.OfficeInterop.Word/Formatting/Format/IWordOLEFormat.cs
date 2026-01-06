//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中 OLE 对象格式的封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOLEFormat : IOfficeObject<IWordOLEFormat, MsWord.OLEFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置指定 OLE 对象、图片或字段的类类型。
    /// </summary>
    string ClassType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定的对象是否显示为图标。
    /// </summary>
    bool DisplayAsIcon { get; set; }

    /// <summary>
    /// 获取或设置存储 OLE 对象图标的程序文件。
    /// </summary>
    string IconName { get; set; }

    /// <summary>
    /// 获取存储 OLE 对象图标的文件路径。
    /// </summary>
    string IconPath { get; }

    /// <summary>
    /// 获取或设置当 DisplayAsIcon 属性为 True 时使用的图标：0（零）对应第一个图标，1 对应第二个图标，依此类推。如果省略此参数，则使用第一个（默认）图标。
    /// </summary>
    int IconIndex { get; set; }

    /// <summary>
    /// 获取或设置显示在 OLE 对象图标下方的文本。
    /// </summary>
    string IconLabel { get; set; }

    /// <summary>
    /// 获取用于标识正在链接的源文件部分的字符串。
    /// 例如，如果源文件是 Microsoft Excel 工作簿，则 Label 属性可能返回 "Workbook1!R3C1:R4C2"（如果 OLE 对象仅包含工作表中的几个单元格）。
    /// </summary>
    string Label { get; }

    /// <summary>
    /// 获取表示指定 OLE 对象顶级接口的对象。
    /// 此属性允许您访问 ActiveX 控件的属性方法，或访问创建 OLE 对象的应用程序的属性方法。OLE 对象必须支持 OLE 自动化才能使此属性正常工作。
    /// </summary>
    object Object { get; }

    /// <summary>
    /// 获取指定 OLE 对象的编程标识符（ProgID）。
    /// </summary>
    string ProgID { get; }

    /// <summary>
    /// 激活指定的对象。
    /// </summary>
    void Activate();

    /// <summary>
    /// 在创建它的应用程序中打开指定的 OLE 对象进行编辑。
    /// </summary>
    void Edit();

    /// <summary>
    /// 打开指定的对象。
    /// </summary>
    void Open();

    /// <summary>
    /// 请求 OLE 对象执行其可用动词之一，即 OLE 对象激活其内容所采取的操作。每个 OLE 对象都支持一组与该对象相关的动词。
    /// </summary>
    /// <param name="verbIndex">可选项。OLE 对象应执行的动词。如果省略此参数，则发送默认动词。如果 OLE 对象不支持请求的动词，将发生错误。可以是任何 WdOLEVerb 常量。</param>
    void DoVerb(WdOLEVerb? verbIndex = null);

    /// <summary>
    /// 将指定的 OLE 对象从一个类转换为另一个类，使您能够在不同的服务器应用程序中编辑该对象，或者更改对象在文档中的显示方式。
    /// </summary>
    /// <param name="classType">可选项。用于激活 OLE 对象的应用程序名称。您可以在“对象”对话框（插入菜单）的“新建”选项卡的“对象类型”框中查看可用应用程序的列表。通过将对象作为内嵌形状插入，然后查看域代码，可以找到 ClassType 字符串。对象的类类型跟在“EMBED”或“LINK”之后。</param>
    /// <param name="displayAsIcon">可选项。True 表示将 OLE 对象显示为图标。默认值为 False。</param>
    /// <param name="iconFileName">可选项。包含要显示的图标的文件。</param>
    /// <param name="iconIndex">可选项。IconFileName 中图标的索引号。指定文件中图标的顺序对应于选中“显示为图标”复选框时“更改图标”对话框（插入菜单，“对象”对话框）中图标的显示顺序。文件中的第一个图标的索引号为 0（零）。如果 IconFileName 中不存在具有给定索引号的图标，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。</param>
    /// <param name="iconLabel">可选项。显示在图标下方的标签（标题）。</param>
    void ConvertTo(object? classType = null, bool? displayAsIcon = null, string? iconFileName = null, int? iconIndex = null, string? iconLabel = null);

    /// <summary>
    /// 设置 Windows 注册表值，该值确定用于激活指定 OLE 对象的默认应用程序。
    /// </summary>
    /// <param name="classType">必需。打开 OLE 对象的应用程序名称。要查看 OLE 对象可以作为激活的对象类型列表，请单击对象，然后打开“转换”对话框（编辑菜单，“对象”子菜单）。通过将对象作为内嵌形状插入，然后查看域代码，可以找到 ClassType 字符串。对象的类类型跟在“EMBED”或“LINK”之后。</param>
    void ActivateAs(string classType);

    /// <summary>
    /// 获取或设置一个值，指示是否保留在 Microsoft Word 中对链接的 OLE 对象所做的格式化，例如链接到 Microsoft Excel 电子表格的表格。
    /// </summary>
    bool PreserveFormattingOnUpdate { get; set; }
}
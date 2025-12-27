//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SmartTagAction 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SmartTagAction 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSmartTagAction : IDisposable
{

    /// <summary>
    /// 获取智能标记动作所在的智能标记对象
    /// 对应 SmartTagAction.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取内部对象所在的Application对象
    /// 对应 Interior.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 执行与单元格上智能标记类型关联的智能标记操作。
    /// </summary>
    void Execute();

    /// <summary>
    /// 获取表示在文档操作任务窗格中显示的智能文档控件类型。
    /// </summary>
    XlSmartTagControlType Type { get; }

    /// <summary>
    /// 获取一个布尔值，表示智能文档控件当前是否显示在文档操作任务窗格中。
    /// </summary>
    bool PresentInPane { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示指定的智能文档帮助文本控件在文档操作任务窗格中是展开还是折叠。
    /// true表示控件已展开，false表示控件已折叠。
    /// </summary>
    bool ExpandHelp { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示智能文档中复选框是否被选中。
    /// </summary>
    bool CheckboxState { get; set; }

    /// <summary>
    /// 获取或设置智能文档文本框控件中的文本。
    /// </summary>
    string TextboxText { get; set; }

    /// <summary>
    /// 获取或设置表示智能文档列表框控件中选定项的索引号。
    /// </summary>
    int ListSelection { get; set; }

    /// <summary>
    /// 获取或设置表示智能文档中单选按钮控件组中选定项的索引号。
    /// </summary>
    int RadioGroupSelection { get; set; }

    /// <summary>
    /// 获取表示在文档操作任务窗格中显示的ActiveX控件。
    /// </summary>
    object ActiveXControl { get; }
}
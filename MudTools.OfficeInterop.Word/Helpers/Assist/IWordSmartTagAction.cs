//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
// SmartTagAction 接口及实现类
public interface IWordSmartTagAction : IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象 [[9]]。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取代表指定对象的父对象的对象 [[18]]。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取智能标记操作的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取代表智能文档控件类型的 WdSmartTagControlType [[15]]。
    /// </summary>
    WdSmartTagControlType Type { get; }

    /// <summary>
    /// 获取或设置一个整数，该整数代表智能文档列表框控件中所选项的索引号 [[7]]。
    /// </summary>
    int ListSelection { get; set; }

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值代表智能文档复选框控件的状态。
    /// </summary>
    bool CheckboxState { get; set; }

    /// <summary>
    /// 获取或设置一个 Object 类型的值，该值代表智能文档 ActiveX 控件的值。
    /// </summary>
    object ActiveXControl { get; set; }

    string TextboxText { get; set; }

    int RadioGroupSelection { get; set; }

    bool ExpandDocumentFragment { get; set; }

    bool ExpandHelp { get; set; }

    bool PresentInPane { get; }

    /// <summary>
    /// 执行指定的智能标记操作 [[14]]。
    /// </summary>
    void Execute();
}
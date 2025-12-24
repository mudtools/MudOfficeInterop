//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 模板（.dotx, .dotm 等）的封装接口。
/// 提供对模板基本属性的访问和修改能力。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTemplate : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取模板的文件全路径（例如：C:\Templates\Normal.dotm）
    /// </summary>
    string? FullName { get; }

    /// <summary>
    /// 获取模板的文件名（包含扩展名，例如：Normal.dotm）
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// 获取模板所在的文件夹路径
    /// </summary>
    string? Path { get; }

    /// <summary>
    /// 获取或设置在指定字符后禁止换行的字符列表
    /// </summary>
    string? NoLineBreakBefore { get; set; }

    /// <summary>
    /// 获取或设置在指定字符前禁止换行的字符列表
    /// </summary>
    string? NoLineBreakAfter { get; set; }

    /// <summary>
    /// 获取或设置不进行拼写检查的语言设置
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取模板中的自动图文集条目集合
    /// </summary>
    IWordAutoTextEntries? AutoTextEntries { get; }

    /// <summary>
    /// 获取模板中的构建基块条目集合
    /// </summary>
    IWordBuildingBlockEntries? BuildingBlockEntries { get; }

    /// <summary>
    /// 获取模板中的构建基块类型集合
    /// </summary>
    IWordBuildingBlockTypes? BuildingBlockTypes { get; }

    /// <summary>
    /// 获取模板中的列表模板集合
    /// </summary>
    IWordListTemplates? ListTemplates { get; }

    /// <summary>
    /// 获取模板中的自定义文档属性集合
    /// </summary>
    IOfficeDocumentProperties? CustomDocumentProperties { get; }

    /// <summary>
    /// 获取模板中的内置文档属性集合
    /// </summary>
    IOfficeDocumentProperties? BuiltInDocumentProperties { get; }

    /// <summary>
    /// 获取或设置是否使用算法调整字距
    /// </summary>
    bool KerningByAlgorithm { get; set; }


    /// <summary>
    /// 获取模板的类型
    /// </summary>
    WdTemplateType Type { get; }

    /// <summary>
    /// 获取或设置文本对齐时的字符间距调整模式
    /// </summary>
    WdJustificationMode JustificationMode { get; set; }

    /// <summary>
    /// 获取或设置远东语言文本的换行控制级别
    /// </summary>
    WdFarEastLineBreakLevel FarEastLineBreakLevel { get; set; }

    /// <summary>
    /// 获取或设置远东语言换行规则的语言标识
    /// </summary>
    WdFarEastLineBreakLanguageID FarEastLineBreakLanguage { get; set; }


    /// <summary>
    /// 将模板作为文档打开
    /// </summary>
    /// <returns>打开的文档对象</returns>
    IWordDocument? OpenAsDocument();

    /// <summary>
    /// 获取或设置模板的打开状态（只读属性，通常不可设置）
    /// </summary>
    bool Saved { get; set; }

    /// <summary>
    /// 保存模板的更改
    /// </summary>
    void Save();

}
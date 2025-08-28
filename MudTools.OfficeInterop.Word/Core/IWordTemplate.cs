namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 模板（.dotx, .dotm 等）的封装接口。
/// 提供对模板基本属性的访问和修改能力。
/// </summary>
public interface IWordTemplate : IDisposable
{
    IWordApplication? Application { get; }

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

    IWordBuildingBlockEntries? BuildingBlockEntries { get; }

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
    IWordDocument OpenAsDocument();

    /// <summary>
    /// 获取或设置模板的打开状态（只读属性，通常不可设置）
    /// </summary>
    bool Saved { get; set; }


    /// <summary>
    /// 保存模板的更改
    /// </summary>
    void Save();

}
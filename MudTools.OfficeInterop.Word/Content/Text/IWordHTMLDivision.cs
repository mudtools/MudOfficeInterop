//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的HTML分节（Division），提供对HTML分节的属性和操作的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordHTMLDivision : IDisposable
{
    /// <summary>
    /// 获取与该HTML分节关联的Word应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该HTML分节的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示该HTML分节所包含文本范围的对象
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取该HTML分节的边框集合
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取或设置该HTML分节的左缩进量
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置该HTML分节的右缩进量
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置该HTML分节前的间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置该HTML分节后的间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取该HTML分节所包含的子HTML分节集合
    /// </summary>
    IWordHTMLDivisions HTMLDivisions { get; }

    /// <summary>
    /// 获取当前HTML分节的父级HTML分节
    /// </summary>
    /// <param name="LevelsUp">向上查找的层级数，默认为1表示直接父级</param>
    /// <returns>指定层级的父HTML分节，如果不存在则返回null</returns>
    IWordHTMLDivision? HTMLDivisionParent(int? LevelsUp);

    /// <summary>
    /// 删除当前HTML分节
    /// </summary>
    void Delete();
}
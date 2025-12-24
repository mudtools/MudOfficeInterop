//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office文档中的脚本集合，提供对集合中脚本的访问、添加和删除操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeScripts : IEnumerable<IOfficeScript?>, IDisposable
{
    /// <summary>
    /// 获取集合中脚本的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的脚本
    /// </summary>
    /// <param name="index">脚本在集合中的索引位置，从0开始</param>
    /// <returns>指定索引位置的脚本对象</returns>
    IOfficeScript? this[int index] { get; }

    /// <summary>
    /// 通过名称获取集合中的脚本
    /// </summary>
    /// <param name="name">脚本的名称</param>
    /// <returns>指定名称的脚本对象</returns>
    IOfficeScript? this[string name] { get; }

    /// <summary>
    /// 删除集合中的所有脚本
    /// </summary>
    void Delete();

    /// <summary>
    /// 向集合中添加新的脚本
    /// </summary>
    /// <param name="anchor">脚本的锚点对象，通常是文档中的一个位置或元素</param>
    /// <param name="location">脚本在文档中的位置，默认为在文档主体中</param>
    /// <param name="language">脚本语言，默认为Visual Basic</param>
    /// <param name="id">脚本的唯一标识符，默认为空字符串</param>
    /// <param name="extended">脚本的扩展属性，默认为空字符串</param>
    /// <param name="scriptText">要添加的脚本代码文本，默认为空字符串</param>
    /// <returns>新添加的脚本对象</returns>
    IOfficeScript? Add(object anchor, MsoScriptLocation location = MsoScriptLocation.msoScriptLocationInBody, MsoScriptLanguage language = MsoScriptLanguage.msoScriptLanguageVisualBasic, string id = "", string extended = "", string scriptText = "");

}
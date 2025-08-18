//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

public interface IWordLanguages : IDisposable, IEnumerable<IWordLanguage>
{
    /// <summary>
    /// 获取语言集合中的语言数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取语言对象（索引从1开始）
    /// </summary>
    /// <param name="index">语言索引</param>
    /// <returns>语言对象</returns>
    IWordLanguage this[int index] { get; }

    /// <summary>
    /// 根据语言ID获取语言对象
    /// </summary>
    /// <param name="languageID">语言ID</param>
    /// <returns>语言对象</returns>
    IWordLanguage GetLanguageByID(int languageID);

    /// <summary>
    /// 检查指定语言ID是否存在于集合中
    /// </summary>
    /// <param name="languageID">语言ID</param>
    /// <returns>是否存在</returns>
    bool Contains(int languageID);
}
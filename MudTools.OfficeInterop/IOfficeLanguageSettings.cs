//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office LanguageSettings 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.LanguageSettings 的安全访问和操作
/// </summary>
public interface IOfficeLanguageSettings : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 LanguageSettings 对象所在的Application对象
    /// 对应 LanguageSettings.Application 属性
    /// </summary>
    object Application { get; }
    #endregion

    #region 语言设置属性
    /// <summary>
    /// 获取指定语言类型的标识符 (ID)
    /// 对应 LanguageSettings.LanguageID 属性 (通过索引器访问)
    /// </summary>
    /// <returns>语言标识符</returns>
    MsoLanguageID GetLanguageID(MsoAppLanguageID languageID);

    /// <summary>
    /// 获取指定语言类型的语言标识符 (ID)
    /// 提供更符合 C# 属性习惯的访问方式
    /// </summary>
    /// <param name="languageID">语言类型 (使用 MsoLanguageID 枚举对应的 int)</param>
    /// <returns>语言标识符</returns>
    MsoLanguageID this[MsoAppLanguageID languageID] { get; }

    /// <summary>
    /// 获取指定应用程序的首选编辑语言是否为指定语言
    /// 对应 LanguageSettings.LanguagePreferredForEditing 属性 (通过索引器访问)
    /// </summary>
    /// <param name="languageID">语言标识符 (使用 MsoLanguageID 枚举对应的 int)</param>
    /// <returns>如果该语言是首选编辑语言则为 true，否则为 false</returns>
    bool GetLanguagePreferredForEditing(MsoLanguageID languageID);
    #endregion   
}
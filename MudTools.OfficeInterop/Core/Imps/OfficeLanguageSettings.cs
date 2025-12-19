//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office LanguageSettings 对象的二次封装实现类
/// 实现 IOfficeLanguageSettings 接口
/// </summary>
internal partial class OfficeLanguageSettings
{

    #region 语言设置属性
    public MsoLanguageID GetLanguageID(MsoAppLanguageID languageID)
    {
        return (MsoLanguageID)_languageSettings.LanguageID[(MsCore.MsoAppLanguageID)languageID];
    }

    public MsoLanguageID this[MsoAppLanguageID languageID]
    {
        get => GetLanguageID(languageID);
    }

    public bool GetLanguagePreferredForEditing(MsoLanguageID languageID)
    {
        return _languageSettings.get_LanguagePreferredForEditing((MsCore.MsoLanguageID)languageID);
    }

    #endregion


}

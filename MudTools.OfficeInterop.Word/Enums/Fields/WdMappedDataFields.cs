//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定可用于邮件合并操作的可映射数据字段类型
/// </summary>
public enum WdMappedDataFields
{

    /// <summary>
    /// 唯一标识符字段
    /// </summary>
    wdUniqueIdentifier = 1,

    /// <summary>
    /// 称谓字段（如先生、女士等）
    /// </summary>
    wdCourtesyTitle,

    /// <summary>
    /// 名字字段
    /// </summary>
    wdFirstName,

    /// <summary>
    /// 中间名字段
    /// </summary>
    wdMiddleName,

    /// <summary>
    /// 姓氏字段
    /// </summary>
    wdLastName,

    /// <summary>
    /// 后缀字段（如 Jr., Sr.等）
    /// </summary>
    wdSuffix,

    /// <summary>
    /// 昵称字段
    /// </summary>
    wdNickname,

    /// <summary>
    /// 职位字段
    /// </summary>
    wdJobTitle,

    /// <summary>
    /// 公司名称字段
    /// </summary>
    wdCompany,

    /// <summary>
    /// 地址行1字段
    /// </summary>
    wdAddress1,

    /// <summary>
    /// 地址行2字段
    /// </summary>
    wdAddress2,

    /// <summary>
    /// 城市字段
    /// </summary>
    wdCity,

    /// <summary>
    /// 州/省字段
    /// </summary>
    wdState,

    /// <summary>
    /// 邮政编码字段
    /// </summary>
    wdPostalCode,

    /// <summary>
    /// 国家/地区字段
    /// </summary>
    wdCountryRegion,

    /// <summary>
    /// 商务电话字段
    /// </summary>
    wdBusinessPhone,

    /// <summary>
    /// 商务传真字段
    /// </summary>
    wdBusinessFax,

    /// <summary>
    /// 家庭电话字段
    /// </summary>
    wdHomePhone,

    /// <summary>
    /// 家庭传真字段
    /// </summary>
    wdHomeFax,

    /// <summary>
    /// 电子邮件地址字段
    /// </summary>
    wdEmailAddress,

    /// <summary>
    /// 网页URL字段
    /// </summary>
    wdWebPageURL,

    /// <summary>
    /// 配偶称谓字段
    /// </summary>
    wdSpouseCourtesyTitle,

    /// <summary>
    /// 配偶名字字段
    /// </summary>
    wdSpouseFirstName,

    /// <summary>
    /// 配偶中间名字段
    /// </summary>
    wdSpouseMiddleName,

    /// <summary>
    /// 配偶姓氏字段
    /// </summary>
    wdSpouseLastName,

    /// <summary>
    /// 配偶昵称字段
    /// </summary>
    wdSpouseNickname,

    /// <summary>
    /// Ruby名字字段（日文发音注音）
    /// </summary>
    wdRubyFirstName,

    /// <summary>
    /// Ruby姓氏字段（日文发音注音）
    /// </summary>
    wdRubyLastName,

    /// <summary>
    /// 地址行3字段
    /// </summary>
    wdAddress3,

    /// <summary>
    /// 部门字段
    /// </summary>
    wdDepartment
}
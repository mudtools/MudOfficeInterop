//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 指定对列进行排序时应用的排序类型
/// </summary>
[Guid("80EB5D8F-AF97-3E3F-9EBD-1B1C645CBCC4")]
public enum WdSortFieldType
{
    /// <summary>
    /// 字母数字顺序
    /// </summary>
    wdSortFieldAlphanumeric,

    /// <summary>
    /// 数字顺序
    /// </summary>
    wdSortFieldNumeric,

    /// <summary>
    /// 日期顺序
    /// </summary>
    wdSortFieldDate,

    /// <summary>
    /// 音节顺序
    /// </summary>
    wdSortFieldSyllable,

    /// <summary>
    /// 日本 JIS 顺序
    /// </summary>
    wdSortFieldJapanJIS,

    /// <summary>
    /// 笔画顺序
    /// </summary>
    wdSortFieldStroke,

    /// <summary>
    /// 韩国 KS 顺序
    /// </summary>
    wdSortFieldKoreaKS
}
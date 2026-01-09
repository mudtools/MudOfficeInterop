//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要创建的数据序列类型。
/// </summary>
public enum XlDataSeriesType
{
    /// <summary>
    /// 根据自动填充设置填充序列。
    /// </summary>
    xlAutoFill = 4,

    /// <summary>
    /// 填充日期值。
    /// </summary>
    xlChronological = 3,

    /// <summary>
    /// 按等比序列扩展数值（例如，“1, 2”将扩展为“4, 8, 16”）。
    /// </summary>
    xlGrowth = 2,

    /// <summary>
    /// 按等差序列扩展数值（例如，“1, 2”将扩展为“3, 4, 5”）。
    /// </summary>
    xlDataSeriesLinear = -4132
}
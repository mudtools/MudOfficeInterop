//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定包含数据透视表区域左上角的数据透视表部分
/// </summary>
public enum XlLocationInTable
{
    /// <summary>
    /// 列标题
    /// </summary>
    xlColumnHeader = -4110,

    /// <summary>
    /// 列项
    /// </summary>
    xlColumnItem = 5,

    /// <summary>
    /// 数据标题
    /// </summary>
    xlDataHeader = 3,

    /// <summary>
    /// 数据项
    /// </summary>
    xlDataItem = 7,

    /// <summary>
    /// 页标题
    /// </summary>
    xlPageHeader = 2,

    /// <summary>
    /// 页项
    /// </summary>
    xlPageItem = 6,

    /// <summary>
    /// 行标题
    /// </summary>
    xlRowHeader = -4153,

    /// <summary>
    /// 行项
    /// </summary>
    xlRowItem = 4,

    /// <summary>
    /// 表体
    /// </summary>
    xlTableBody = 8
}
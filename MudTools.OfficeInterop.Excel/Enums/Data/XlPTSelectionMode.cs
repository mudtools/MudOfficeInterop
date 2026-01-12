//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据透视表的结构化选择模式
/// 用于控制在数据透视表中可以选择的内容类型
/// </summary>
public enum XlPTSelectionMode
{
    /// <summary>
    /// 仅选择标签
    /// </summary>
    xlLabelOnly = 1,

    /// <summary>
    /// 选择数据和标签
    /// </summary>
    xlDataAndLabel = 0,

    /// <summary>
    /// 仅选择数据
    /// </summary>
    xlDataOnly = 2,

    /// <summary>
    /// 选择原点（Origin）
    /// </summary>
    xlOrigin = 3,

    /// <summary>
    /// 选择按钮
    /// </summary>
    xlButton = 15,

    /// <summary>
    /// 选择空白区域
    /// </summary>
    xlBlanks = 4,

    /// <summary>
    /// 选择第一行
    /// 可与其他常量组合使用
    /// </summary>
    xlFirstRow = 256
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定具有条件格式的数据条单元格范围的轴位置
/// </summary>
public enum XlDataBarAxisPosition
{
    /// <summary>
    /// 根据范围内最小负值与最大正值的比例显示轴。正值从左到右显示，负值从右到左显示。当所有值均为正值或均为负值时，不显示轴。
    /// </summary>
    xlDataBarAxisAutomatic,

    /// <summary>
    /// 无论范围内的值集如何，始终在单元格中点显示轴。正值从左到右显示，负值从右到左显示。
    /// </summary>
    xlDataBarAxisMidpoint,

    /// <summary>
    /// 不显示轴，正值和负值均从左到右方向显示。
    /// </summary>
    xlDataBarAxisNone
}
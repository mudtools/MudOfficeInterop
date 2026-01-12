//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 方向枚举
/// 用于指定在单元格区域中的移动方向
/// </summary>
public enum XlDirection
{
    /// <summary>
    /// 向下方向
    /// 在单元格区域中向下移动
    /// </summary>
    xlDown = -4121,

    /// <summary>
    /// 向左方向
    /// 在单元格区域中向左移动
    /// </summary>
    xlToLeft = -4159,

    /// <summary>
    /// 向右方向
    /// 在单元格区域中向右移动
    /// </summary>
    xlToRight = -4161,

    /// <summary>
    /// 向上方向
    /// 在单元格区域中向上移动
    /// </summary>
    xlUp = -4162
}
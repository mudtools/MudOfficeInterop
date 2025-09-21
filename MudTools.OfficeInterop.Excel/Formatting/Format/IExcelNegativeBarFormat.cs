//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel NegativeBarFormat 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.NegativeBarFormat 的安全访问和操作
/// </summary>
public interface IExcelNegativeBarFormat : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取负值数据条格式对象的父对象 (通常是 DataBar)
    /// 对应 NegativeBarFormat.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取负值数据条格式对象所在的Application对象
    /// 对应 NegativeBarFormat.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置负值数据条的颜色类型
    /// 对应 NegativeBarFormat.ColorType 属性
    /// </summary>
    int ColorType { get; set; } // Use int for XlDataBarNegativeColorType

    /// <summary>
    /// 获取或设置负值数据条的边框颜色类型
    /// 对应 NegativeBarFormat.BorderColorType 属性
    /// </summary>
    int BorderColorType { get; set; } // Use int for XlDataBarNegativeColorType

    /// <summary>
    /// 获取或设置负值数据条的填充颜色
    /// 对应 NegativeBarFormat.Color 属性
    /// </summary>
    int Color { get; }

    /// <summary>
    /// 获取或设置负值数据条的边框颜色
    /// 对应 NegativeBarFormat.BorderColor 属性
    /// </summary>
    int BorderColor { get; }
    #endregion
}

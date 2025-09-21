//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Databar 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Databar 的安全访问和操作
/// </summary>
public interface IExcelDataBar : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据条对象的父对象 (通常是 FormatCondition)
    /// 对应 Databar.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据条对象所在的Application对象
    /// 对应 Databar.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置最小条件值
    /// 对应 Databar.MinPoint 属性
    /// </summary>
    IExcelConditionValue MinPoint { get; }

    /// <summary>
    /// 获取或设置最大条件值
    /// 对应 Databar.MaxPoint 属性
    /// </summary>
    IExcelConditionValue MaxPoint { get; }

    /// <summary>
    /// 获取或设置数据条的方向
    /// 对应 Databar.Direction 属性
    /// </summary>
    int Direction { get; set; }

    /// <summary>
    /// 获取或设置数据条的图形条显示
    /// 对应 Databar.BarFillType 属性
    /// </summary>
    XlDataBarFillType BarFillType { get; set; }

    /// <summary>
    /// 获取或设置是否显示数据条的边框
    /// 对应 Databar.BarBorder 属性的设置
    /// </summary>
    bool ShowBarOnly { get; set; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取数据条的边框对象
    /// </summary>
    IExcelDataBarBorder Borders { get; }

    IExcelNegativeBarFormat NegativeBarFormat { get; }

    /// <summary>
    /// 获取数据条的字体对象
    /// </summary>
    string Formula { get; set; }
    #endregion
}

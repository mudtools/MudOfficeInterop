//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ColorScaleCriterion 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ColorScaleCriterion 的安全访问和操作
/// ColorScaleCriterion 对象代表颜色刻度条件格式中的一个特定条件（例如，最小值、中间值、最大值）。
/// </summary>
public interface IExcelColorScaleCriterion : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取颜色刻度条件对象所在的Application对象
    /// 对应 ColorScaleCriterion.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取颜色刻度条件的索引（在 ColorScaleCriteria 集合中）
    /// 对应 ColorScaleCriterion.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置颜色刻度条件的类型
    /// 对应 ColorScaleCriterion.Type 属性
    /// </summary>
    XlConditionValueTypes Type { get; set; }

    /// <summary>
    /// 获取或设置颜色刻度条件的值
    /// 对应 ColorScaleCriterion.Value 属性
    /// </summary>
    object Value { get; set; } // 可以是数字、百分比、公式字符串等

    /// <summary>
    /// 获取或设置颜色刻度条件对应的颜色
    /// 对应 ColorScaleCriterion.FormatColor 属性 或相关颜色属性
    /// </summary>
    int Color { get; set; }
    #endregion
}
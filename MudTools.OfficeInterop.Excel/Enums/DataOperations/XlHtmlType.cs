//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定 Microsoft Excel 将指定项保存为网页时生成的 HTML 类型，以及该项是静态还是交互式
/// </summary>
public enum XlHtmlType
{
    /// <summary>
    /// 使用静态（非交互式）HTML，仅用于查看
    /// </summary>
    xlHtmlStatic,

    /// <summary>
    /// 使用电子表格组件
    /// </summary>
    xlHtmlCalc,

    /// <summary>
    /// 使用数据透视表组件
    /// </summary>
    xlHtmlList,

    /// <summary>
    /// 使用图表组件
    /// </summary>
    xlHtmlChart
}
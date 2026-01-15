//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel数据透视表公式接口，用于操作Excel中的数据透视表公式
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotFormula : IOfficeObject<IExcelPivotFormula, MsExcel.PivotFormula>, IDisposable
{
    /// <summary>
    /// 获取图表标题的父对象
    /// 对应 ChartTitle.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图表标题所在的 Application 对象
    /// 对应 ChartTitle.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置公式在集合中的索引位置
    /// </summary>
    int Index { get; set; }

    /// <summary>
    /// 获取或设置数据透视表公式的公式字符串
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置公式的值
    /// </summary>
    string Value { get; set; }

    /// <summary>
    /// 获取或设置标准公式字符串
    /// </summary>
    string StandardFormula { get; set; }

    /// <summary>
    /// 删除当前数据透视表公式
    /// </summary>
    void Delete();
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel数据透视表筛选器接口，用于操作Excel中的数据透视表筛选器
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotFilter : IOfficeObject<IExcelPivotFilter>, IDisposable
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
    /// 获取或设置筛选器在筛选器列表中的顺序
    /// </summary>
    int Order { get; set; }

    /// <summary>
    /// 获取筛选器的类型
    /// </summary>
    XlPivotFilterType FilterType { get; }

    /// <summary>
    /// 获取筛选器的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取筛选器的描述信息
    /// </summary>
    string Description { get; }

    /// <summary>
    /// 获取筛选器是否处于激活状态
    /// </summary>
    bool Active { get; }

    /// <summary>
    /// 获取与筛选器关联的数据透视字段
    /// </summary>
    IExcelPivotField? PivotField { get; }

    /// <summary>
    /// 获取筛选器的数据字段
    /// </summary>
    IExcelPivotField? DataField { get; }

    /// <summary>
    /// 获取筛选器的数据立方体字段
    /// </summary>
    IExcelCubeField? DataCubeField { get; }

    /// <summary>
    /// 获取成员属性字段
    /// </summary>
    IExcelPivotField? MemberPropertyField { get; }

    /// <summary>
    /// 获取筛选器的第一个值
    /// </summary>
    object Value1 { get; }

    /// <summary>
    /// 获取筛选器的第二个值
    /// </summary>
    object Value2 { get; }

    /// <summary>
    /// 获取是否为成员属性筛选器
    /// </summary>
    bool IsMemberPropertyFilter { get; }

    /// <summary>
    /// 获取或设置是否按整天进行筛选
    /// </summary>
    bool WholeDayFilter { get; set; }

    /// <summary>
    /// 删除当前筛选器
    /// </summary>
    void Delete();
}
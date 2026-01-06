//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示Excel中Top10条件格式规则的接口
/// 该接口用于定义和操作Excel中的Top10条件格式规则，可以设置显示前N个或后N个值的格式
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTop10 : IOfficeObject<IExcelTop10>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置条件格式规则的优先级
    /// 优先级较高的规则会先于优先级较低的规则进行计算和应用
    /// </summary>
    int Priority { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当条件为真时是否停止评估其他条件格式规则
    /// 如果设置为true，则当此条件满足时，不再评估后续的条件格式规则
    /// </summary>
    bool StopIfTrue { get; set; }

    /// <summary>
    /// 获取条件格式规则应用的范围
    /// </summary>
    IExcelRange? AppliesTo { get; }

    /// <summary>
    /// 获取条件格式规则的内部区域格式
    /// 可用于设置满足条件的单元格的背景色等内部格式
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取条件格式规则的边框格式
    /// 可用于设置满足条件的单元格的边框样式
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 获取条件格式规则的字体格式
    /// 可用于设置满足条件的单元格的字体样式
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置是显示前N个还是后N个值
    /// </summary>
    /// <value>
    /// xlTop10Top表示显示前N个值，xlTop10Bottom表示显示后N个值
    /// </value>
    XlTopBottom TopBottom { get; set; }

    /// <summary>
    /// 获取或设置条件格式规则的作用范围类型
    /// </summary>
    XlPivotConditionScope ScopeType { get; set; }

    /// <summary>
    /// 获取或设置计算的范围
    /// </summary>
    XlCalcFor CalcFor { get; set; }

    /// <summary>
    /// 获取条件格式规则的类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取一个值，指示条件格式是否与数据透视表相关
    /// </summary>
    bool PTCondition { get; }

    /// <summary>
    /// 获取或设置要显示的项目数（排名）
    /// 例如，设置为10表示显示前10项或后10项
    /// </summary>
    int Rank { get; set; }

    /// <summary>
    /// 获取或设置排名是否基于百分比
    /// 如果设置为true，则Rank属性表示百分比；否则表示实际项数
    /// </summary>
    bool Percent { get; set; }

    /// <summary>
    /// 获取或设置满足条件的单元格的数字格式
    /// </summary>
    object NumberFormat { get; set; }

    /// <summary>
    /// 将此条件格式规则设置为最高优先级
    /// </summary>
    void SetFirstPriority();

    /// <summary>
    /// 将此条件格式规则设置为最低优先级
    /// </summary>
    void SetLastPriority();

    /// <summary>
    /// 删除此条件格式规则
    /// </summary>
    void Delete();

    /// <summary>
    /// 修改条件格式规则应用的范围
    /// </summary>
    /// <param name="Range">新的应用范围</param>
    void ModifyAppliesToRange(IExcelRange Range);
}
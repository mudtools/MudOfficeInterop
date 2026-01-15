//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示一个图标集条件格式规则
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelIconSetCondition : IOfficeObject<IExcelIconSetCondition, MsExcel.IconSetCondition>, IDisposable
{
    /// <summary>
    /// 获取条件格式规则的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件格式规则所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取一个值，指示条件格式规则是否为最高优先级
    /// </summary>
    bool StopIfTrue { get; }

    /// <summary>
    /// 获取条件格式规则的类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取一个值，指示条件格式是否与数据透视表相关
    /// </summary>
    bool PTCondition { get; }

    /// <summary>
    /// 获取或设置数据透视表条件格式的作用范围类型
    /// </summary>
    XlPivotConditionScope ScopeType { get; set; }

    /// <summary>
    /// 获取条件格式规则应用的单元格区域
    /// </summary>
    IExcelRange? AppliesTo { get; }

    /// <summary>
    /// 获取或设置图标集
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelIconSets? IconSet { get; set; }

    /// <summary>
    /// 获取图标集的条件标准集合
    /// </summary>
    IExcelIconCriteria? IconCriteria { get; }

    /// <summary>
    /// 获取一个值，指示是否只显示图标而不显示单元格值
    /// </summary>
    bool ShowIconOnly { get; }

    /// <summary>
    /// 获取或设置条件格式规则的优先级
    /// </summary>
    int Priority { get; set; }

    /// <summary>
    /// 获取或设置是否反转图标顺序
    /// </summary>
    bool ReverseOrder { get; set; }

    /// <summary>
    /// 获取或设置是否使用百分位数值
    /// </summary>
    bool PercentileValues { get; set; }

    /// <summary>
    /// 获取或设置确定将对图标集应用的值的公式
    /// </summary>
    string Formula { get; set; }


    /// <summary>
    /// 修改条件格式规则应用的单元格区域
    /// </summary>
    /// <param name="Range">要应用条件格式的单元格区域</param>
    void ModifyAppliesToRange(IExcelRange Range);

    /// <summary>
    /// 将条件格式规则设置为最高优先级
    /// </summary>
    void SetFirstPriority();

    /// <summary>
    /// 将条件格式规则设置为最低优先级
    /// </summary>
    void SetLastPriority();

    /// <summary>
    /// 删除条件格式规则
    /// </summary>
    void Delete();

}
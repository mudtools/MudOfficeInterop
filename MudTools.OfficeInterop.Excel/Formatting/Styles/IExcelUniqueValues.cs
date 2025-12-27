//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的唯一值条件格式规则接口
/// 该接口用于定义和操作唯一值或重复值的条件格式规则
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelUniqueValues : IOfficeObject<IExcelUniqueValues>, IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置条件格式规则的优先级
    /// 优先级数字越小，规则越先执行
    /// </summary>
    int Priority { get; set; }


    /// <summary>
    /// 获取或设置是否在条件为真时停止执行后续规则
    /// </summary>
    bool StopIfTrue { get; set; }

    /// <summary>
    /// 获取条件格式规则的类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取一个值，指示条件格式是否与数据透视表相关
    /// </summary>
    bool PTCondition { get; }

    /// <summary>
    /// 获取或设置条件格式规则的数字格式
    /// </summary>
    object NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置数据透视表条件格式的作用范围类型
    /// </summary>
    XlPivotConditionScope ScopeType { get; set; }

    /// <summary>
    /// 获取条件格式规则的应用范围
    /// </summary>
    IExcelRange? AppliesTo { get; }

    /// <summary>
    /// 获取条件格式规则的内部填充格式对象
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取条件格式规则的边框格式对象
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 获取条件格式规则的字体格式对象
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置是显示唯一值还是重复值
    /// </summary>
    XlDupeUnique DupeUnique { get; set; }

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
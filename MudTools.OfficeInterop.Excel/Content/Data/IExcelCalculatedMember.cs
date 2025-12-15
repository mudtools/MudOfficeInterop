//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 数据模型中的计算成员（Calculated Member）接口。
/// 计算成员是在数据透视表或多维数据集中定义的成员，它不是直接来自数据源，
/// 而是通过对其他成员执行计算得到的。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCalculatedMember : IDisposable
{
    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取计算成员的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取计算成员的公式。
    /// </summary>
    string Formula { get; }

    /// <summary>
    /// 获取计算成员的数据源名称。
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 获取计算成员的求解顺序。
    /// 求解顺序决定了在计算过程中何时计算该成员，数字越大优先级越高。
    /// </summary>
    int SolveOrder { get; }

    /// <summary>
    /// 获取一个值，指示计算成员是否有效。
    /// </summary>
    bool IsValid { get; }

    /// <summary>
    /// 获取计算成员的类型。
    /// </summary>
    XlCalculatedMemberType Type { get; }

    /// <summary>
    /// 获取一个值，指示计算成员是否为动态成员。
    /// </summary>
    bool Dynamic { get; }

    /// <summary>
    /// 获取计算成员的显示文件夹路径。
    /// </summary>
    string DisplayFolder { get; }

    /// <summary>
    /// 获取或设置一个值，指示在层次结构化时是否区分唯一值。
    /// </summary>
    bool HierarchizeDistinct { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否将层次结构扁平化。
    /// </summary>
    bool FlattenHierarchies { get; }

    /// <summary>
    /// 获取计算成员关联的度量值组。
    /// </summary>
    string MeasureGroup { get; }

    /// <summary>
    /// 获取计算成员的父层次结构。
    /// </summary>
    string ParentHierarchy { get; }

    /// <summary>
    /// 获取计算成员的父成员。
    /// </summary>
    string ParentMember { get; }

    /// <summary>
    /// 获取计算成员的数字格式类型。
    /// </summary>
    XlCalcMemNumberFormatType NumberFormat { get; }

    /// <summary>
    /// 删除计算成员。
    /// </summary>
    void Delete();
}
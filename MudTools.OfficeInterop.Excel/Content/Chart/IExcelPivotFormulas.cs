//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 数据透视表公式集合的接口。
/// 此接口封装了与 Microsoft Excel 数据透视表公式相关的功能，允许访问和操作数据透视表中的计算字段公式。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotFormulas : IOfficeObject<IExcelPivotFormulas>, IEnumerable<IExcelPivotFormula?>, IDisposable
{
    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取集合中数据透视表公式的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的数据透视表公式项。
    /// </summary>
    /// <param name="index">要获取的数据透视表公式在集合中的从1开始的索引位置。</param>
    /// <returns>指定索引位置的数据透视表公式对象。</returns>
    IExcelPivotFormula? this[int index] { get; }

    /// <summary>
    /// 通过名称获取集合中的数据透视表公式项。
    /// </summary>
    /// <param name="name">要获取的数据透视表公式的名称。</param>
    /// <returns>具有指定名称的数据透视表公式对象。</returns>
    IExcelPivotFormula? this[string name] { get; }

    /// <summary>
    /// 向数据透视表公式集合中添加新的公式。
    /// </summary>
    /// <param name="Formula">要添加到数据透视表中的公式字符串。</param>
    /// <param name="useStandardFormula">可选参数，指示是否使用标准公式格式。如果为 null，则使用默认设置。</param>
    /// <returns>新添加的数据透视表公式对象。</returns>
    IExcelPivotFormula? Add(string Formula, bool? useStandardFormula);
}
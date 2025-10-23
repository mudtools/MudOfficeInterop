

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示数据透视表中的所有计算字段的集合。
/// 此接口是对 Microsoft.Office.Interop.Excel.CalculatedFields COM 对象的封装。
/// 注意：集合中的每个计算字段实际上是一个 <see cref="IExcelPivotField"/> 对象。
/// </summary>
public interface IExcelCalculatedFields : IEnumerable<IExcelPivotField>, IDisposable
{
    /// <summary>
    /// 获取集合中的计算字段总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引或名称的计算字段。
    /// 索引从 1 开始，也可以使用字段的名称（字符串）作为索引。
    /// </summary>
    /// <param name="index">计算字段的索引（从1开始）或名称。</param>
    /// <returns>指定的 <see cref="IExcelPivotField"/> 对象。</returns>
    IExcelPivotField? this[object index] { get; }

    /// <summary>
    /// 获取该对象的父对象（通常是数据透视表）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 在数据透视表中创建一个新的计算字段。
    /// </summary>
    /// <param name="name">新计算字段的名称。</param>
    /// <param name="formula">新计算字段的公式。</param>
    /// <param name="useStandardFormula">
    /// 如果为 true，则假定公式使用标准英语（美国）格式。
    /// 如果为 false，则假定公式采用本地化格式。
    /// </param>
    /// <returns>新创建的 <see cref="IExcelPivotField"/> 对象。</returns>
    IExcelPivotField? Add(string name, string formula, bool useStandardFormula = true);
}

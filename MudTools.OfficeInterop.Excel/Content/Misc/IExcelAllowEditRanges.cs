namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示所有 AllowEditRange 对象的集合，这些对象表示受保护工作表中可以编辑的单元格。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelAllowEditRanges : IEnumerable<IExcelAllowEditRange?>, IOfficeObject<IExcelAllowEditRanges, MsExcel.AllowEditRanges>, IDisposable
{
    /// <summary>
    /// 获取集合中的对象数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的 AllowEditRange 对象（索引器）。
    /// </summary>
    /// <param name="index">对象的名称或索引号。</param>
    /// <returns>指定索引处的 AllowEditRange 对象。</returns>
    IExcelAllowEditRange? this[int index] { get; }

    /// <summary>
    /// 获取指定索引处的 AllowEditRange 对象（索引器）。
    /// </summary>
    /// <param name="name">对象的名称或索引号。</param>
    /// <returns>指定索引处的 AllowEditRange 对象。</returns>
    IExcelAllowEditRange? this[string name] { get; }

    /// <summary>
    /// 添加可在受保护工作表上编辑的区域。
    /// </summary>
    /// <param name="title">必需。区域的标题。</param>
    /// <param name="range">必需。Range 对象，表示允许编辑的区域。</param>
    /// <param name="password">可选项。区域的密码。</param>
    /// <returns>新创建的 AllowEditRange 对象。</returns>
    IExcelAllowEditRange Add(string title, IExcelRange range, string? password = null);

}
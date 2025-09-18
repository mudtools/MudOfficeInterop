
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示切片器中所有可筛选项的集合，支持遍历、索引和名称访问。
/// </summary>
public interface IExcelSlicerItems : IEnumerable<IExcelSlicerItem>, IDisposable
{
    /// <summary>
    /// 获取集合中切片器项的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器项。
    /// </summary>
    /// <param name="indexOrName">项索引（int）或名称（string）</param>
    /// <returns>对应的切片器项对象</returns>
    IExcelSlicerItem this[object indexOrName] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Slicer 或 SlicerCache）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 选中所有项（显示所有数据）。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 取消选中所有项（隐藏所有数据，除非设置为允许空筛选）。
    /// </summary>
    void UnselectAll();
}
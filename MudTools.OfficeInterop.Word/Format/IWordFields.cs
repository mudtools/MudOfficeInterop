namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Fields 的接口，用于操作Word域集合。
/// </summary>
public interface IWordFields : IEnumerable<IWordField>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取域集合中的域数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取域（从1开始）。
    /// </summary>
    IWordField this[int index] { get; }

    /// <summary>
    /// 根据范围获取域。
    /// </summary>
    IWordField this[IWordRange range] { get; }

    /// <summary>
    /// 获取域集合的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 更新所有域。
    /// </summary>
    /// <returns>成功更新的域数量。</returns>
    int Update();

    /// <summary>
    /// 取消所有域的链接。
    /// </summary>
    void Unlink();

    /// <summary>
    /// 删除所有域。
    /// </summary>
    void Delete();

    /// <summary>
    /// 根据域类型获取域列表。
    /// </summary>
    /// <param name="fieldType">域类型。</param>
    /// <returns>域列表。</returns>
    List<IWordField> GetFieldsByType(WdFieldType fieldType);

    /// <summary>
    /// 检查是否包含指定类型的域。
    /// </summary>
    /// <param name="fieldType">域类型。</param>
    /// <returns>是否包含。</returns>
    bool ContainsType(WdFieldType fieldType);

    /// <summary>
    /// 获取指定类型的域数量。
    /// </summary>
    /// <param name="fieldType">域类型。</param>
    /// <returns>域数量。</returns>
    int GetCountByType(WdFieldType fieldType);

    /// <summary>
    /// 获取所有域类型。
    /// </summary>
    /// <returns>域类型列表。</returns>
    List<WdFieldType> GetAllFieldTypes();

    /// <summary>
    /// 添加新的域。
    /// </summary>
    /// <param name="range">插入范围。</param>
    /// <param name="type">域类型。</param>
    /// <param name="text">域文本。</param>
    /// <param name="preserveFormatting">是否保持格式。</param>
    /// <returns>新添加的域。</returns>
    IWordField Add(IWordRange range, WdFieldType type, string text = "", bool preserveFormatting = true);

    /// <summary>
    /// 获取所有日期域。
    /// </summary>
    /// <returns>日期域列表。</returns>
    List<IWordField> GetDateFields();

    /// <summary>
    /// 获取所有页码域。
    /// </summary>
    /// <returns>页码域列表。</returns>
    List<IWordField> GetPageFields();

    /// <summary>
    /// 获取所有目录域。
    /// </summary>
    /// <returns>目录域列表。</returns>
    List<IWordField> GetTOCFields();

    /// <summary>
    /// 获取所有链接域。
    /// </summary>
    /// <returns>链接域列表。</returns>
    List<IWordField> GetLinkedFields();

    /// <summary>
    /// 刷新域集合。
    /// </summary>
    void Refresh();

    /// <summary>
    /// 获取域集合中第一个域。
    /// </summary>
    IWordField FirstField { get; }

    /// <summary>
    /// 获取域集合中最后一个域。
    /// </summary>
    IWordField LastField { get; }

    /// <summary>
    /// 根据索引范围获取域子集。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    /// <returns>域子集。</returns>
    List<IWordField> GetFieldsInRange(int startIndex, int endIndex);

    /// <summary>
    /// 清理无效域。
    /// </summary>
    /// <returns>清理的域数量。</returns>
    int CleanupInvalidFields();
}
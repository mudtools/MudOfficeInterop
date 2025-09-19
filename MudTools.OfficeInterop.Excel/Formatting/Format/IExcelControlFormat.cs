
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 表单控件（如复选框、组合框、列表框、滚动条等）的格式设置接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ControlFormat
/// 用于管理控件的列表项、当前值、范围、多选等属性。
/// </summary>
public interface IExcelControlFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置控件的当前选中项索引（从 1 开始）。
    /// 对于多选控件，返回第一个选中项。
    /// </summary>
    int Value { get; set; }

    /// <summary>
    /// 获取或设置控件允许的最小值（适用于滚动条、微调项等）。
    /// </summary>
    int Min { get; set; }

    /// <summary>
    /// 获取或设置控件允许的最大值（适用于滚动条、微调项等）。
    /// </summary>
    int Max { get; set; }

    /// <summary>
    /// 获取或设置控件是否允许多选（适用于列表框）。
    /// </summary>
    bool MultiSelect { get; set; }

    /// <summary>
    /// 获取控件中列表项的总数。
    /// </summary>
    int ListCount { get; }

    /// <summary>
    /// 获取或设置控件中当前选中项的索引（从1开始）。
    /// 对于多选控件，可以设置或获取主选项的索引位置。
    /// </summary>
    int ListIndex { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示控件的文本是否被锁定。
    /// 当设置为 true 时，用户无法编辑控件中的文本内容。
    /// </summary>
    bool LockedText { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示当工作表打印时控件对象是否会被打印。
    /// 设置为 true 表示打印工作表时包含该对象，false 表示不打印该对象。
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 获取或设置用户在控件上使用小幅度操作时的改变值（例如使用鼠标滚轮或方向键）。
    /// 该值通常小于 Increment 值，用于精细调整控件值。
    /// </summary>
    int SmallChange { get; set; }

    /// <summary>
    /// 获取或设置与控件关联的数据源区域（用于动态填充列表项）。
    /// </summary>
    string ListFillRange { get; set; }

    /// <summary>
    /// 获取或设置与控件值绑定的单元格（控件值变化时自动写入该单元格）。
    /// </summary>
    string LinkedCell { get; set; }

    /// <summary>
    /// 向控件列表中添加一个新项。
    /// </summary>
    /// <param name="text">要添加的文本。</param>
    /// <param name="index">插入位置（从1开始），0=追加到末尾。</param>
    void AddItem(string text, int index = 0);

    /// <summary>
    /// 从控件列表中移除指定索引的项。
    /// </summary>
    /// <param name="index">要删除的项索引（从1开始）。</param>
    void RemoveItem(int index);

    /// <summary>
    /// 清空控件中的所有列表项。
    /// </summary>
    void RemoveAllItems();

    /// <summary>
    /// 获取指定索引项的文本内容。
    /// </summary>
    /// <param name="index">项索引（从1开始）。</param>
    /// <returns>项文本，若无效则返回空字符串。</returns>
    string GetItemText(int index);
}
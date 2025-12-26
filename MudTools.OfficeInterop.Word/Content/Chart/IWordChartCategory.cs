namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表分类的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartCategory : IOfficeObject<IWordChartCategory>, IDisposable
{
    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取分类名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置分类的筛选状态。
    /// </summary>
    bool IsFiltered { get; set; }


}

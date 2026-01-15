namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel自定义属性集合的接口，用于管理和操作Excel工作簿的自定义文档属性集合
/// 该接口封装了COM对象，提供对Excel自定义属性集合的访问和操作能力
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCustomProperties : IOfficeObject<IExcelCustomProperties, MsExcel.CustomProperties>, IEnumerable<IExcelCustomProperty?>, IDisposable
{
    /// <summary>
    /// 获取自定义属性集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取自定义属性集合所属的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取自定义属性的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取自定义属性
    /// </summary>
    /// <param name="index">自定义属性索引</param>
    /// <returns>自定义属性对象</returns>
    IExcelCustomProperty? this[int index] { get; }

    /// <summary>
    /// 通过名称获取自定义属性
    /// </summary>
    /// <param name="name">自定义属性名称</param>
    /// <returns>自定义属性对象</returns>
    IExcelCustomProperty? this[string name] { get; }

    /// <summary>
    /// 向集合中添加新的自定义属性
    /// </summary>
    /// <param name="name">自定义属性的名称</param>
    /// <param name="value">自定义属性的值</param>
    /// <returns>新创建的自定义属性对象</returns>
    IExcelCustomProperty? Add(string name, object value);
}
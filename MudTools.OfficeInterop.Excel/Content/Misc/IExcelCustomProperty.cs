namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel自定义属性的接口，用于管理和操作Excel工作簿的自定义文档属性
/// 该接口封装了COM对象，提供对Excel自定义属性的访问和操作能力
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCustomProperty : IOfficeObject<IExcelCustomProperty, MsExcel.CustomProperty>, IDisposable
{
    /// <summary>
    /// 获取自定义属性的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取自定义属性所属的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置自定义属性的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置自定义属性的值
    /// </summary>
    object Value { get; set; }

    /// <summary>
    /// 删除自定义属性
    /// </summary>
    void Delete();
}
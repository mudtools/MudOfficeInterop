namespace MudTools.OfficeInterop.Excel;

public interface IExcelError : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    IExcelApplication Application { get; }


    /// <summary>
    /// 获取错误值
    /// </summary>
    object Value { get; }

    /// <summary>
    /// 获取错误是否为忽略错误
    /// </summary>
    bool Ignore { get; set; }
}
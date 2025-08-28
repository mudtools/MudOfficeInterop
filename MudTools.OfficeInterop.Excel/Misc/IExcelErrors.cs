namespace MudTools.OfficeInterop.Excel;

using System;

/// <summary>
/// Excel Errors COM组件二次封装
/// </summary>
public interface IExcelErrors : IDisposable
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
    /// 通过索引获取错误
    /// </summary>
    /// <param name="index">错误索引</param>
    /// <returns>错误对象</returns>
    IExcelError this[object index] { get; }
}
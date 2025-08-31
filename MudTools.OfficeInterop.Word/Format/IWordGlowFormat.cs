namespace MudTools.OfficeInterop.Word;

using System;

/// <summary>
/// 表示 Word 发光效果格式的封装接口。
/// </summary>
public interface IWordGlowFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置发光颜色。
    /// </summary>
    IWordColorFormat? Color { get; }

    /// <summary>
    /// 获取或设置发光半径（磅）。
    /// </summary>
    float Radius { get; set; }


    float Transparency { get; set; }
}

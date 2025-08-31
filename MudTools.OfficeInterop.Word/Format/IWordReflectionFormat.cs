namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 倒影效果格式的封装接口。
/// </summary>
public interface IWordReflectionFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置倒影类型。
    /// </summary>
    MsoReflectionType Type { get; set; }

    /// <summary>
    /// 获取或设置倒影透明度（0-100）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置倒影大小。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置倒影偏移量。
    /// </summary>
    float Offset { get; set; }

    /// <summary>
    /// 获取或设置倒影模糊度。
    /// </summary>
    float Blur { get; set; }
}
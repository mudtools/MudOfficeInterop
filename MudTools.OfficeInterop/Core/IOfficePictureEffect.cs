namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureEffect 的接口，用于操作图片效果。
/// </summary>
public interface IOfficePictureEffect : IDisposable
{
    /// <summary>
    /// 获取或设置效果的类型。
    /// </summary>
    MsoPictureEffectType Type { get; }

    /// <summary>
    /// 获取或设置效果的位置。
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取或设置效果是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取效果的参数集合。
    /// </summary>
    IOfficeEffectParameters EffectParameters { get; }

    /// <summary>
    /// 删除此图片效果。
    /// </summary>
    void Delete();
}
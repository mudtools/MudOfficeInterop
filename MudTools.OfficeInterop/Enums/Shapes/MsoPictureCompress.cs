namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图片压缩选项，用于控制Office应用程序中的图片压缩行为
/// </summary>
public enum MsoPictureCompress
{
    /// <summary>
    /// 使用文档默认设置进行图片压缩
    /// </summary>
    msoPictureCompressDocDefault = -1,

    /// <summary>
    /// 不压缩图片
    /// </summary>
    msoPictureCompressFalse,

    /// <summary>
    /// 压缩图片
    /// </summary>
    msoPictureCompressTrue
}
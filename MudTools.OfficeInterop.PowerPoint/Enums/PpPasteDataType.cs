//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定粘贴操作的数据类型。
/// </summary>
public enum PpPasteDataType
{
    /// <summary>
    /// 默认粘贴数据类型。
    /// </summary>
    ppPasteDefault,

    /// <summary>
    /// 位图粘贴数据类型。
    /// </summary>
    ppPasteBitmap,

    /// <summary>
    /// 增强型图元文件粘贴数据类型。
    /// </summary>
    ppPasteEnhancedMetafile,

    /// <summary>
    /// 图元文件图片粘贴数据类型。
    /// </summary>
    ppPasteMetafilePicture,

    /// <summary>
    /// GIF 粘贴数据类型。
    /// </summary>
    ppPasteGIF,

    /// <summary>
    /// JPG 粘贴数据类型。
    /// </summary>
    ppPasteJPG,

    /// <summary>
    /// PNG 粘贴数据类型。
    /// </summary>
    ppPastePNG,

    /// <summary>
    /// 文本粘贴数据类型。
    /// </summary>
    ppPasteText,

    /// <summary>
    /// HTML 粘贴数据类型。
    /// </summary>
    ppPasteHTML,

    /// <summary>
    /// RTF 粘贴数据类型。
    /// </summary>
    ppPasteRTF,

    /// <summary>
    /// OLE 对象粘贴数据类型。
    /// </summary>
    ppPasteOLEObject,

    /// <summary>
    /// 形状粘贴数据类型。
    /// </summary>
    ppPasteShape
}
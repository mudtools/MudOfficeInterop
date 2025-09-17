//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在Word文档中粘贴数据时的数据类型
/// </summary>
public enum WdPasteDataType
{
    /// <summary>
    /// OLE对象格式
    /// </summary>
    wdPasteOLEObject = 0,
    
    /// <summary>
    /// RTF格式（Rich Text Format）
    /// </summary>
    wdPasteRTF = 1,
    
    /// <summary>
    /// 纯文本格式
    /// </summary>
    wdPasteText = 2,
    
    /// <summary>
    /// 图元文件图片格式
    /// </summary>
    wdPasteMetafilePicture = 3,
    
    /// <summary>
    /// 位图格式
    /// </summary>
    wdPasteBitmap = 4,
    
    /// <summary>
    /// 设备无关位图格式
    /// </summary>
    wdPasteDeviceIndependentBitmap = 5,
    
    /// <summary>
    /// 超链接格式
    /// </summary>
    wdPasteHyperlink = 7,
    
    /// <summary>
    /// 形状对象格式
    /// </summary>
    wdPasteShape = 8,
    
    /// <summary>
    /// 增强型图元文件格式
    /// </summary>
    wdPasteEnhancedMetafile = 9,
    
    /// <summary>
    /// HTML格式
    /// </summary>
    wdPasteHTML = 10
}
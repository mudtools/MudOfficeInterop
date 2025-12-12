//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状类型的枚举，用于定义Office应用程序中各种对象的形状类型
/// </summary>
public enum MsoShapeType
{
    /// <summary>
    /// 混合形状类型
    /// </summary>
    msoShapeTypeMixed = -2,

    /// <summary>
    /// 自动形状（如矩形、圆形等基本图形）
    /// </summary>
    msoAutoShape = 1,

    /// <summary>
    /// 标注形状
    /// </summary>
    msoCallout = 2,

    /// <summary>
    /// 图表
    /// </summary>
    msoChart = 3,

    /// <summary>
    /// 批注
    /// </summary>
    msoComment = 4,

    /// <summary>
    /// 自由曲线
    /// </summary>
    msoFreeform = 5,

    /// <summary>
    /// 组合形状
    /// </summary>
    msoGroup = 6,

    /// <summary>
    /// 嵌入的OLE对象
    /// </summary>
    msoEmbeddedOLEObject = 7,

    /// <summary>
    /// 表单控件
    /// </summary>
    msoFormControl = 8,

    /// <summary>
    /// 直线
    /// </summary>
    msoLine = 9,

    /// <summary>
    /// 链接的OLE对象
    /// </summary>
    msoLinkedOLEObject = 10,

    /// <summary>
    /// 链接的图片
    /// </summary>
    msoLinkedPicture = 11,

    /// <summary>
    /// OLE控件对象
    /// </summary>
    msoOLEControlObject = 12,

    /// <summary>
    /// 图片
    /// </summary>
    msoPicture = 13,

    /// <summary>
    /// 占位符
    /// </summary>
    msoPlaceholder = 14,

    /// <summary>
    /// 文本效果
    /// </summary>
    msoTextEffect = 15,

    /// <summary>
    /// 媒体对象
    /// </summary>
    msoMedia = 16,

    /// <summary>
    /// 文本框
    /// </summary>
    msoTextBox = 17,

    /// <summary>
    /// 脚本锚点
    /// </summary>
    msoScriptAnchor = 18,

    /// <summary>
    /// 表格
    /// </summary>
    msoTable = 19,

    /// <summary>
    /// 墨迹对象
    /// </summary>
    msoInk = 20,

    /// <summary>
    /// 墨迹批注
    /// </summary>
    msoInkComment = 21,

    /// <summary>
    /// SmartArt图形
    /// </summary>
    msoSmartArt = 22,

    /// <summary>
    /// 切片器
    /// </summary>
    msoSlicer = 23,

    /// <summary>
    /// 日期时间选择器
    /// </summary>
    msoDateTimePicker = 24,

    /// <summary>
    /// 网络视频
    /// </summary>
    msoWebVideo = 25
}
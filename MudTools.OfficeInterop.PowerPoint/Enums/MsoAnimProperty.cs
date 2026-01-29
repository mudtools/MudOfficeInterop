//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定可用于动画的属性。
/// </summary>
public enum MsoAnimProperty
{
    /// <summary>
    /// 无属性动画。
    /// </summary>
    msoAnimNone = 0,

    /// <summary>
    /// X坐标动画属性。
    /// </summary>
    msoAnimX = 1,

    /// <summary>
    /// Y坐标动画属性。
    /// </summary>
    msoAnimY = 2,

    /// <summary>
    /// 宽度动画属性。
    /// </summary>
    msoAnimWidth = 3,

    /// <summary>
    /// 高度动画属性。
    /// </summary>
    msoAnimHeight = 4,

    /// <summary>
    /// 不透明度动画属性。
    /// </summary>
    msoAnimOpacity = 5,

    /// <summary>
    /// 旋转角度动画属性。
    /// </summary>
    msoAnimRotation = 6,

    /// <summary>
    /// 颜色动画属性。
    /// </summary>
    msoAnimColor = 7,

    /// <summary>
    /// 可见性动画属性。
    /// </summary>
    msoAnimVisibility = 8,

    /// <summary>
    /// 文本字体加粗动画属性。
    /// </summary>
    msoAnimTextFontBold = 100,

    /// <summary>
    /// 文本字体颜色动画属性。
    /// </summary>
    msoAnimTextFontColor = 101,

    /// <summary>
    /// 文本字体浮雕效果动画属性。
    /// </summary>
    msoAnimTextFontEmboss = 102,

    /// <summary>
    /// 文本字体斜体动画属性。
    /// </summary>
    msoAnimTextFontItalic = 103,

    /// <summary>
    /// 文本字体名称动画属性。
    /// </summary>
    msoAnimTextFontName = 104,

    /// <summary>
    /// 文本字体阴影动画属性。
    /// </summary>
    msoAnimTextFontShadow = 105,

    /// <summary>
    /// 文本字体大小动画属性。
    /// </summary>
    msoAnimTextFontSize = 106,

    /// <summary>
    /// 文本字体下标动画属性。
    /// </summary>
    msoAnimTextFontSubscript = 107,

    /// <summary>
    /// 文本字体上标动画属性。
    /// </summary>
    msoAnimTextFontSuperscript = 108,

    /// <summary>
    /// 文本字体下划线动画属性。
    /// </summary>
    msoAnimTextFontUnderline = 109,

    /// <summary>
    /// 文本字体删除线动画属性。
    /// </summary>
    msoAnimTextFontStrikeThrough = 110,

    /// <summary>
    /// 文本项目符号字符动画属性。
    /// </summary>
    msoAnimTextBulletCharacter = 111,

    /// <summary>
    /// 文本项目符号字体名称动画属性。
    /// </summary>
    msoAnimTextBulletFontName = 112,

    /// <summary>
    /// 文本项目符号编号动画属性。
    /// </summary>
    msoAnimTextBulletNumber = 113,

    /// <summary>
    /// 文本项目符号颜色动画属性。
    /// </summary>
    msoAnimTextBulletColor = 114,

    /// <summary>
    /// 文本项目符号相对大小动画属性。
    /// </summary>
    msoAnimTextBulletRelativeSize = 115,

    /// <summary>
    /// 文本项目符号样式动画属性。
    /// </summary>
    msoAnimTextBulletStyle = 116,

    /// <summary>
    /// 文本项目符号类型动画属性。
    /// </summary>
    msoAnimTextBulletType = 117,

    /// <summary>
    /// 形状图片对比度动画属性。
    /// </summary>
    msoAnimShapePictureContrast = 1000,

    /// <summary>
    /// 形状图片亮度动画属性。
    /// </summary>
    msoAnimShapePictureBrightness = 1001,

    /// <summary>
    /// 形状图片伽马值动画属性。
    /// </summary>
    msoAnimShapePictureGamma = 1002,

    /// <summary>
    /// 形状图片灰度动画属性。
    /// </summary>
    msoAnimShapePictureGrayscale = 1003,

    /// <summary>
    /// 形状填充启用动画属性。
    /// </summary>
    msoAnimShapeFillOn = 1004,

    /// <summary>
    /// 形状填充颜色动画属性。
    /// </summary>
    msoAnimShapeFillColor = 1005,

    /// <summary>
    /// 形状填充不透明度动画属性。
    /// </summary>
    msoAnimShapeFillOpacity = 1006,

    /// <summary>
    /// 形状填充背景色动画属性。
    /// </summary>
    msoAnimShapeFillBackColor = 1007,

    /// <summary>
    /// 形状线条启用动画属性。
    /// </summary>
    msoAnimShapeLineOn = 1008,

    /// <summary>
    /// 形状线条颜色动画属性。
    /// </summary>
    msoAnimShapeLineColor = 1009,

    /// <summary>
    /// 形状阴影启用动画属性。
    /// </summary>
    msoAnimShapeShadowOn = 1010,

    /// <summary>
    /// 形状阴影类型动画属性。
    /// </summary>
    msoAnimShapeShadowType = 1011,

    /// <summary>
    /// 形状阴影颜色动画属性。
    /// </summary>
    msoAnimShapeShadowColor = 1012,

    /// <summary>
    /// 形状阴影不透明度动画属性。
    /// </summary>
    msoAnimShapeShadowOpacity = 1013,

    /// <summary>
    /// 形状阴影X偏移量动画属性。
    /// </summary>
    msoAnimShapeShadowOffsetX = 1014,

    /// <summary>
    /// 形状阴影Y偏移量动画属性。
    /// </summary>
    msoAnimShapeShadowOffsetY = 1015
}
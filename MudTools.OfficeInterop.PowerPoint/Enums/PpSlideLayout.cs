//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 幻灯片布局类型枚举
/// </summary>
public enum PpSlideLayout
{
    /// <summary>
    /// 混合布局
    /// </summary>
    ppLayoutMixed = -2,
    
    /// <summary>
    /// 标题布局
    /// </summary>
    ppLayoutTitle = 1,
    
    /// <summary>
    /// 文本布局
    /// </summary>
    ppLayoutText = 2,
    
    /// <summary>
    /// 两栏文本布局
    /// </summary>
    ppLayoutTwoColumnText = 3,
    
    /// <summary>
    /// 表格布局
    /// </summary>
    ppLayoutTable = 4,
    
    /// <summary>
    /// 文本和图表布局
    /// </summary>
    ppLayoutTextAndChart = 5,
    
    /// <summary>
    /// 图表和文本布局
    /// </summary>
    ppLayoutChartAndText = 6,
    
    /// <summary>
    /// 组织结构图布局
    /// </summary>
    ppLayoutOrgchart = 7,
    
    /// <summary>
    /// 图表布局
    /// </summary>
    ppLayoutChart = 8,
    
    /// <summary>
    /// 文本和剪贴画布局
    /// </summary>
    ppLayoutTextAndClipart = 9,
    
    /// <summary>
    /// 剪贴画和文本布局
    /// </summary>
    ppLayoutClipartAndText = 10,
    
    /// <summary>
    /// 仅标题布局
    /// </summary>
    ppLayoutTitleOnly = 11,
    
    /// <summary>
    /// 空白布局
    /// </summary>
    ppLayoutBlank = 12,
    
    /// <summary>
    /// 文本和对象布局
    /// </summary>
    ppLayoutTextAndObject = 13,
    
    /// <summary>
    /// 对象和文本布局
    /// </summary>
    ppLayoutObjectAndText = 14,
    
    /// <summary>
    /// 大对象布局
    /// </summary>
    ppLayoutLargeObject = 15,
    
    /// <summary>
    /// 对象布局
    /// </summary>
    ppLayoutObject = 16,
    
    /// <summary>
    /// 文本和媒体剪辑布局
    /// </summary>
    ppLayoutTextAndMediaClip = 17,
    
    /// <summary>
    /// 媒体剪辑和文本布局
    /// </summary>
    ppLayoutMediaClipAndText = 18,
    
    /// <summary>
    /// 对象在文本上方布局
    /// </summary>
    ppLayoutObjectOverText = 19,
    
    /// <summary>
    /// 文本在对象上方布局
    /// </summary>
    ppLayoutTextOverObject = 20,
    
    /// <summary>
    /// 文本和两个对象布局
    /// </summary>
    ppLayoutTextAndTwoObjects = 21,
    
    /// <summary>
    /// 两个对象和文本布局
    /// </summary>
    ppLayoutTwoObjectsAndText = 22,
    
    /// <summary>
    /// 两个对象在文本上方布局
    /// </summary>
    ppLayoutTwoObjectsOverText = 23,
    
    /// <summary>
    /// 四个对象布局
    /// </summary>
    ppLayoutFourObjects = 24,
    
    /// <summary>
    /// 垂直文本布局
    /// </summary>
    ppLayoutVerticalText = 25,
    
    /// <summary>
    /// 剪贴画和垂直文本布局
    /// </summary>
    ppLayoutClipArtAndVerticalText = 26,
    
    /// <summary>
    /// 垂直标题和文本布局
    /// </summary>
    ppLayoutVerticalTitleAndText = 27,
    
    /// <summary>
    /// 垂直标题和文本在图表上方布局
    /// </summary>
    ppLayoutVerticalTitleAndTextOverChart = 28,
    
    /// <summary>
    /// 两个对象布局
    /// </summary>
    ppLayoutTwoObjects = 29,
    
    /// <summary>
    /// 对象和两个对象布局
    /// </summary>
    ppLayoutObjectAndTwoObjects = 30,
    
    /// <summary>
    /// 两个对象和对象布局
    /// </summary>
    ppLayoutTwoObjectsAndObject = 31,
    
    /// <summary>
    /// 自定义布局
    /// </summary>
    ppLayoutCustom = 32,
    
    /// <summary>
    /// 节标题布局
    /// </summary>
    ppLayoutSectionHeader = 33,
    
    /// <summary>
    /// 两个文本和两个对象布局
    /// </summary>
    ppLayoutTwoTextAndTwoObjects = 34,
    
    /// <summary>
    /// 标题、对象和题注布局
    /// </summary>
    ppLayoutTitleObjectAndCaption = 35,
    
    /// <summary>
    /// 带题注的图片布局
    /// </summary>
    ppLayoutPictureWithCaption = 36,
    
    /// <summary>
    /// 垂直对象布局
    /// </summary>
    ppLayoutVerticalObject = 37,
    
    /// <summary>
    /// 对象和按钮布局
    /// </summary>
    ppLayoutObjectAndButton = 38
}

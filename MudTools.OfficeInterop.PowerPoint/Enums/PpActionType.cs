//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 动作类型枚举
/// </summary>
public enum PpActionType
{
    /// <summary>
    /// 无动作
    /// </summary>
    ppActionNone = 0,

    /// <summary>
    /// 跳转到下一张幻灯片
    /// </summary>
    ppActionNextSlide = 1,

    /// <summary>
    /// 跳转到上一张幻灯片
    /// </summary>
    ppActionPreviousSlide = 2,

    /// <summary>
    /// 跳转到第一张幻灯片
    /// </summary>
    ppActionFirstSlide = 3,

    /// <summary>
    /// 跳转到最后一张幻灯片
    /// </summary>
    ppActionLastSlide = 4,

    /// <summary>
    /// 跳转到最近查看的幻灯片
    /// </summary>
    ppActionLastSlideViewed = 5,

    /// <summary>
    /// 结束放映
    /// </summary>
    ppActionEndShow = 6,

    /// <summary>
    /// 超链接动作
    /// </summary>
    ppActionHyperlink = 7,

    /// <summary>
    /// 运行宏
    /// </summary>
    ppActionRunMacro = 8,

    /// <summary>
    /// 运行程序
    /// </summary>
    ppActionRunProgram = 9,

    /// <summary>
    /// 命名幻灯片放映
    /// </summary>
    ppActionNamedSlideShow = 10,

    /// <summary>
    /// OLE动作
    /// </summary>
    ppActionOLEVerb = 11,

    /// <summary>
    /// 播放媒体
    /// </summary>
    ppActionPlay = 12
}
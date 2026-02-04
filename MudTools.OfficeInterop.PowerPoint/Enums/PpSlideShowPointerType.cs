//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定幻灯片放映时使用的指针类型。
/// </summary>
public enum PpSlideShowPointerType
{
    /// <summary>
    /// 无指针。
    /// </summary>
    ppSlideShowPointerNone,

    /// <summary>
    /// 箭头指针。
    /// </summary>
    ppSlideShowPointerArrow,

    /// <summary>
    /// 笔形指针（可用于标注）。
    /// </summary>
    ppSlideShowPointerPen,

    /// <summary>
    /// 始终隐藏指针。
    /// </summary>
    ppSlideShowPointerAlwaysHidden,

    /// <summary>
    /// 自动箭头指针（鼠标移动时显示，静止时隐藏）。
    /// </summary>
    ppSlideShowPointerAutoArrow,

    /// <summary>
    /// 橡皮擦指针（用于擦除笔迹）。
    /// </summary>
    ppSlideShowPointerEraser
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 视图接口，用于操作 PowerPoint 视图
/// </summary>
public interface IPowerPointView : IDisposable
{
    /// <summary>
    /// 获取视图类型
    /// </summary>
    PpViewType Type { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取当前幻灯片
    /// </summary>
    IPowerPointSlide Slide { get; }

    /// <summary>
    /// 获取当前幻灯片索引
    /// </summary>
    int SlideIndex { get; }

    /// <summary>
    /// 获取或设置缩放比例
    /// </summary>
    int Zoom { get; set; }

    /// <summary>
    /// 获取选中的形状范围
    /// </summary>
    IPowerPointShapeRange Selection { get; }

    /// <summary>
    /// 转到指定幻灯片
    /// </summary>
    /// <param name="slideIndex">幻灯片索引</param>
    void GoToSlide(int slideIndex);

    /// <summary>
    /// 开始幻灯片放映
    /// </summary>
    /// <param name="fromSlide">起始幻灯片索引</param>
    /// <param name="toSlide">结束幻灯片索引</param>
    /// <returns>幻灯片放映窗口对象</returns>
    IPowerPointSlideShowWindow StartSlideShow(int fromSlide = 1, int toSlide = -1);

    /// <summary>
    /// 激活当前视图
    /// </summary>
    void Activate();
}
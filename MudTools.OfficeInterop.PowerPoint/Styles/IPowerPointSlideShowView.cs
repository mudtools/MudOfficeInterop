//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 幻灯片放映视图接口
/// </summary>
public interface IPowerPointSlideShowView : IDisposable
{
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
    /// 获取幻灯片放映状态
    /// </summary>
    PpSlideShowState State { get; }

    /// <summary>
    /// 转到指定幻灯片
    /// </summary>
    /// <param name="slideIndex">幻灯片索引</param>
    void GoToSlide(int slideIndex);

    /// <summary>
    /// 转到下一张幻灯片
    /// </summary>
    void NextSlide();

    /// <summary>
    /// 转到上一张幻灯片
    /// </summary>
    void PreviousSlide();

    /// <summary>
    /// 转到第一张幻灯片
    /// </summary>
    void FirstSlide();

    /// <summary>
    /// 转到最后一张幻灯片
    /// </summary>
    void LastSlide();


    /// <summary>
    /// 结束幻灯片放映
    /// </summary>
    void End();
}

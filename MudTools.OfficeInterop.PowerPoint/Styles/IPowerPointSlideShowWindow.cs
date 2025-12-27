//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 幻灯片放映窗口接口
/// </summary>
public interface IPowerPointSlideShowWindow : IDisposable
{
    /// <summary>
    /// 获取或设置窗口高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置窗口宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置窗口左边缘位置
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置窗口上边缘位置
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片放映视图
    /// </summary>
    IPowerPointSlideShowView View { get; }

    /// <summary>
    /// 获取幻灯片放映设置
    /// </summary>
    IPowerPointSlideShowSettings Settings { get; }

    /// <summary>
    /// 获取幻灯片放映状态
    /// </summary>
    PpSlideShowState State { get; }

    /// <summary>
    /// 激活窗口
    /// </summary>
    void Activate();

    /// <summary>
    /// 暂停幻灯片放映
    /// </summary>
    void Pause();

    /// <summary>
    /// 恢复幻灯片放映
    /// </summary>
    void Resume();

    /// <summary>
    /// 切换到黑屏
    /// </summary>
    void BlackScreen();

    /// <summary>
    /// 切换到白屏
    /// </summary>
    void WhiteScreen();
}
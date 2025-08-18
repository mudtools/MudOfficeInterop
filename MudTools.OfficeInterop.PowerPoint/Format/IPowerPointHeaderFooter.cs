//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 页眉页脚项接口
/// </summary>
public interface IPowerPointHeaderFooter : IDisposable
{
    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置文本
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置格式
    /// </summary>
    int Format { get; set; }

    /// <summary>
    /// 获取或设置是否使用格式
    /// </summary>
    bool UseFormat { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置位置
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取或设置对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置字体
    /// </summary>
    IPowerPointFont Font { get; }

    /// <summary>
    /// 更新页眉页脚
    /// </summary>
    void Update();

    /// <summary>
    /// 应用到指定幻灯片
    /// </summary>
    /// <param name="slide">目标幻灯片</param>
    void ApplyTo(IPowerPointSlide slide);

    /// <summary>
    /// 设置文本和格式
    /// </summary>
    /// <param name="text">文本内容</param>
    /// <param name="format">格式</param>
    /// <param name="useFormat">是否使用格式</param>
    void SetTextAndFormat(string text, int format = 0, bool useFormat = false);

    /// <summary>
    /// 获取页眉页脚项信息
    /// </summary>
    /// <returns>页眉页脚项信息字符串</returns>
    string GetHeaderFooterInfo();
}
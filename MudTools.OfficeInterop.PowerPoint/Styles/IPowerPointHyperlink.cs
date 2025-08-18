//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 超链接接口
/// </summary>
public interface IPowerPointHyperlink : IDisposable
{
    /// <summary>
    /// 获取或设置超链接地址
    /// </summary>
    string Address { get; set; }

    /// <summary>
    /// 获取或设置子地址
    /// </summary>
    string SubAddress { get; set; }

    /// <summary>
    /// 获取或设置显示文本
    /// </summary>
    string TextToDisplay { get; set; }

    /// <summary>
    /// 获取或设置屏幕提示
    /// </summary>
    string ScreenTip { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取超链接类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取超链接是否有效
    /// </summary>
    bool IsValid { get; }

    /// <summary>
    /// 跟随超链接
    /// </summary>
    void Follow();

    /// <summary>
    /// 删除超链接
    /// </summary>
    void Delete();

    /// <summary>
    /// 编辑超链接
    /// </summary>
    /// <param name="newAddress">新地址</param>
    /// <param name="newSubAddress">新子地址</param>
    /// <param name="newTextToDisplay">新显示文本</param>
    void Edit(string newAddress = null, string newSubAddress = null, string newTextToDisplay = null);

    /// <summary>
    /// 复制超链接
    /// </summary>
    /// <returns>复制的超链接对象</returns>
    IPowerPointHyperlink Duplicate();

    /// <summary>
    /// 应用超链接到指定范围
    /// </summary>
    /// <param name="range">目标范围</param>
    void ApplyTo(object range);

    /// <summary>
    /// 验证超链接
    /// </summary>
    /// <returns>是否有效</returns>
    bool Validate();

    /// <summary>
    /// 获取超链接信息
    /// </summary>
    /// <returns>超链接信息字符串</returns>
    string GetHyperlinkInfo();
}

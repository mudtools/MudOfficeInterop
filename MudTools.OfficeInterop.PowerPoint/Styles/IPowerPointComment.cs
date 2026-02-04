//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片中的注释。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointComment : IDisposable
{
    /// <summary>
    /// 获取创建此注释的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此注释的父对象。
    /// </summary>
    /// <value>表示此注释父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取注释的作者姓名。
    /// </summary>
    /// <value>表示作者姓名的字符串。</value>
    string? Author { get; }

    /// <summary>
    /// 获取注释的作者首字母缩写。
    /// </summary>
    /// <value>表示作者首字母缩写的字符串。</value>
    string? AuthorInitials { get; }

    /// <summary>
    /// 获取注释的文本内容。
    /// </summary>
    /// <value>表示注释文本的字符串。</value>
    string? Text { get; }

    /// <summary>
    /// 获取注释的创建日期和时间。
    /// </summary>
    /// <value>表示创建日期时间的 <see cref="System.DateTime"/> 值。</value>
    DateTime DateTime { get; }

    /// <summary>
    /// 获取作者在演示文稿中的索引。
    /// </summary>
    /// <value>表示作者索引的整数值。</value>
    int AuthorIndex { get; }

    /// <summary>
    /// 获取注释在幻灯片上的左边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示左边缘位置的浮点数。</value>
    float Left { get; }

    /// <summary>
    /// 获取注释在幻灯片上的上边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示上边缘位置的浮点数。</value>
    float Top { get; }

    /// <summary>
    /// 删除此注释。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取注释提供程序的标识符。
    /// </summary>
    /// <value>表示提供程序标识符的字符串。</value>
    string? ProviderID { get; }

    /// <summary>
    /// 获取注释用户的标识符。
    /// </summary>
    /// <value>表示用户标识符的字符串。</value>
    string? UserID { get; }

    /// <summary>
    /// 获取注释创建时的时区偏差（以分钟为单位）。
    /// </summary>
    /// <value>表示时区偏差的整数值。</value>
    int TimeZoneBias { get; }

    /// <summary>
    /// 获取此注释的回复集合。
    /// </summary>
    /// <value>表示回复集合的 <see cref="IPowerPointComments"/> 对象。</value>
    IPowerPointComments? Replies { get; }

    /// <summary>
    /// 获取一个值，指示此注释是否已折叠。
    /// </summary>
    /// <value>指示注释是否已折叠的布尔值。</value>
    bool Collapsed { get; }
}
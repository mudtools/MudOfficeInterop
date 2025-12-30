//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示邮件标签的全局邮件标签首选项。
/// <para>注：使用 Application.MailingLabel 属性可返回 MailingLabel 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCustomLabel : IOfficeObject<IWordCustomLabel>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取标签在标签集合中的索引位置
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置标签的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置标签的顶部边距
    /// </summary>
    float TopMargin { get; set; }

    /// <summary>
    /// 获取或设置标签的侧边边距
    /// </summary>
    float SideMargin { get; set; }

    /// <summary>
    /// 获取或设置标签的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置标签的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置标签的垂直间距
    /// </summary>
    float VerticalPitch { get; set; }

    /// <summary>
    /// 获取或设置标签的水平间距
    /// </summary>
    float HorizontalPitch { get; set; }

    /// <summary>
    /// 获取或设置一行中标签的个数
    /// </summary>
    int NumberAcross { get; set; }

    /// <summary>
    /// 获取或设置一列中标签的个数
    /// </summary>
    int NumberDown { get; set; }

    /// <summary>
    /// 获取一个值，该值指示标签是否为点阵标签
    /// </summary>
    bool DotMatrix { get; }

    /// <summary>
    /// 获取或设置自定义标签的页面大小
    /// </summary>
    WdCustomLabelPageSize PageSize { get; set; }

    /// <summary>
    /// 获取一个值，该值指示标签是否有效
    /// </summary>
    bool Valid { get; }

    /// <summary>
    /// 删除当前自定义标签
    /// </summary>
    void Delete();
}
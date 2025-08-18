//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 段落格式接口
/// </summary>
public interface IPowerPointParagraphFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置基线对齐方式
    /// </summary>
    int BaseLineAlignment { get; set; }

    /// <summary>
    /// 获取或设置段落间距控制
    /// </summary>
    int SpaceWithin { get; set; }

    /// <summary>
    /// 获取或设置段落间距类型
    /// </summary>
    int SpaceWithinType { get; set; }

    /// <summary>
    /// 获取或设置是否保持在一起
    /// </summary>
    bool KeepTogether { get; set; }

    /// <summary>
    /// 获取或设置是否保持与下一段在一起
    /// </summary>
    bool KeepWithNext { get; set; }

    /// <summary>
    /// 获取或设置页面分段
    /// </summary>
    bool PageBreakBefore { get; set; }

    /// <summary>
    /// 获取或设置大纲级别
    /// </summary>
    int OutlineLevel { get; set; }

    /// <summary>
    /// 复制段落格式
    /// </summary>
    /// <returns>复制的段落格式对象</returns>
    IPowerPointParagraphFormat Duplicate();

    /// <summary>
    /// 应用段落格式到指定文本范围
    /// </summary>
    /// <param name="textRange">目标文本范围</param>
    void ApplyTo(IPowerPointTextRange textRange);

    /// <summary>
    /// 重置段落格式为默认值
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置段落间距
    /// </summary>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    void SetSpacing(float spaceBefore, float spaceAfter);

    /// <summary>
    /// 设置对齐方式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    void SetAlignment(int alignment);
}

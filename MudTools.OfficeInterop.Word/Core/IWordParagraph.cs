//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档段落接口
/// </summary>
public interface IWordParagraph : IDisposable
{
    /// <summary>
    /// 获取段落范围
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取或设置段落文本
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置段落对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置首行缩进
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置左缩进
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置右缩进
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置行距
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 删除段落
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制段落
    /// </summary>
    void Copy();

    /// <summary>
    /// 选择段落
    /// </summary>
    void Select();
}

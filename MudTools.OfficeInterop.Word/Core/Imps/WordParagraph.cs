//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档段落实现类
/// </summary>
internal class WordParagraph : IWordParagraph
{
    private readonly MsWord.Paragraph _paragraph;
    private bool _disposedValue;
    private IWordRange _range;

    /// <summary>
    /// 获取段落范围
    /// </summary>
    public IWordRange Range
    {
        get
        {
            if (_range == null)
            {
                _range = new WordRange(_paragraph.Range);
            }
            return _range;
        }
    }

    /// <summary>
    /// 获取或设置段落文本
    /// </summary>
    public string Text
    {
        get => _paragraph.Range.Text;
        set => _paragraph.Range.Text = value ?? string.Empty;
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _paragraph.Parent;

    /// <summary>
    /// 获取或设置段落对齐方式
    /// </summary>
    public int Alignment
    {
        get => (int)_paragraph.Alignment;
        set => _paragraph.Alignment = (MsWord.WdParagraphAlignment)value;
    }

    /// <summary>
    /// 获取或设置首行缩进
    /// </summary>
    public float FirstLineIndent
    {
        get => _paragraph.FirstLineIndent;
        set => _paragraph.FirstLineIndent = value;
    }

    /// <summary>
    /// 获取或设置左缩进
    /// </summary>
    public float LeftIndent
    {
        get => _paragraph.LeftIndent;
        set => _paragraph.LeftIndent = value;
    }

    /// <summary>
    /// 获取或设置右缩进
    /// </summary>
    public float RightIndent
    {
        get => _paragraph.RightIndent;
        set => _paragraph.RightIndent = value;
    }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    public float SpaceBefore
    {
        get => _paragraph.SpaceBefore;
        set => _paragraph.SpaceBefore = value;
    }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    public float SpaceAfter
    {
        get => _paragraph.SpaceAfter;
        set => _paragraph.SpaceAfter = value;
    }

    /// <summary>
    /// 获取或设置行距
    /// </summary>
    public float LineSpacing
    {
        get => _paragraph.LineSpacing;
        set => _paragraph.LineSpacing = value;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="paragraph">COM Paragraph 对象</param>
    internal WordParagraph(MsWord.Paragraph paragraph)
    {
        _paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
        _disposedValue = false;
    }

    /// <summary>
    /// 删除段落
    /// </summary>
    public void Delete()
    {
        try
        {
            _paragraph.Range.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete paragraph.", ex);
        }
    }

    /// <summary>
    /// 复制段落
    /// </summary>
    public void Copy()
    {
        try
        {
            _paragraph.Range.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy paragraph.", ex);
        }
    }

    /// <summary>
    /// 选择段落
    /// </summary>
    public void Select()
    {
        try
        {
            _paragraph.Range.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select paragraph.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _range?.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

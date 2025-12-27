//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 段落格式实现类
/// </summary>
internal class PowerPointParagraphFormat : IPowerPointParagraphFormat
{
    private readonly MsPowerPoint.ParagraphFormat _paragraphFormat;
    private bool _disposedValue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _paragraphFormat.Parent;

    /// <summary>
    /// 获取或设置对齐方式
    /// </summary>
    public int Alignment
    {
        get => (int)_paragraphFormat.Alignment;
        set => _paragraphFormat.Alignment = (MsPowerPoint.PpParagraphAlignment)value;
    }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    public float SpaceBefore
    {
        get => _paragraphFormat.SpaceBefore;
        set => _paragraphFormat.SpaceBefore = value;
    }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    public float SpaceAfter
    {
        get => _paragraphFormat.SpaceAfter;
        set => _paragraphFormat.SpaceAfter = value;
    }

    /// <summary>
    /// 获取或设置基线对齐方式
    /// </summary>
    public int BaseLineAlignment
    {
        get => (int)_paragraphFormat.BaseLineAlignment;
        set => _paragraphFormat.BaseLineAlignment = (MsPowerPoint.PpBaselineAlignment)value;
    }

    /// <summary>
    /// 获取或设置段落间距控制
    /// </summary>
    public int SpaceWithin
    {
        get => 0; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置段落间距类型
    /// </summary>
    public int SpaceWithinType
    {
        get => 0; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置是否保持在一起
    /// </summary>
    public bool KeepTogether
    {
        get => false; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置是否保持与下一段在一起
    /// </summary>
    public bool KeepWithNext
    {
        get => false; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置页面分段
    /// </summary>
    public bool PageBreakBefore
    {
        get => false; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置大纲级别
    /// </summary>
    public int OutlineLevel
    {
        get => 0; // PowerPoint 中没有直接对应属性
        set { /* 不实现 */ }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="paragraphFormat">COM ParagraphFormat 对象</param>
    internal PowerPointParagraphFormat(MsPowerPoint.ParagraphFormat paragraphFormat)
    {
        _paragraphFormat = paragraphFormat ?? throw new ArgumentNullException(nameof(paragraphFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 复制段落格式
    /// </summary>
    /// <returns>复制的段落格式对象</returns>
    public IPowerPointParagraphFormat Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法，返回新的实例
            throw new NotImplementedException("Duplicating paragraph format is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate paragraph format.", ex);
        }
    }

    /// <summary>
    /// 应用段落格式到指定文本范围
    /// </summary>
    /// <param name="textRange">目标文本范围</param>
    public void ApplyTo(IPowerPointTextRange textRange)
    {
        if (textRange == null)
            throw new ArgumentNullException(nameof(textRange));

        try
        {
            // 这需要具体的实现来应用格式到文本范围
            throw new NotImplementedException("Applying paragraph format to text range is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply paragraph format to text range.", ex);
        }
    }

    /// <summary>
    /// 重置段落格式为默认值
    /// </summary>
    public void Reset()
    {
        try
        {
            _paragraphFormat.Alignment = MsPowerPoint.PpParagraphAlignment.ppAlignLeft;
            _paragraphFormat.SpaceBefore = 0;
            _paragraphFormat.SpaceAfter = 0;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset paragraph format.", ex);
        }
    }

    /// <summary>
    /// 设置段落间距
    /// </summary>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    public void SetSpacing(float spaceBefore, float spaceAfter)
    {
        try
        {
            _paragraphFormat.SpaceBefore = spaceBefore;
            _paragraphFormat.SpaceAfter = spaceAfter;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set paragraph spacing.", ex);
        }
    }


    /// <summary>
    /// 设置对齐方式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    public void SetAlignment(int alignment)
    {
        try
        {
            _paragraphFormat.Alignment = (MsPowerPoint.PpParagraphAlignment)alignment;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set paragraph alignment.", ex);
        }
    }


    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
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

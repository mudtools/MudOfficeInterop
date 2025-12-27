//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.PowerPoint;

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 页眉页脚项实现类
/// </summary>
internal class PowerPointHeaderFooter : IPowerPointHeaderFooter
{
    private readonly MsPowerPoint.HeaderFooter _headerFooter;
    private bool _disposedValue;
    private IPowerPointFont _font;

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _headerFooter?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headerFooter != null)
                _headerFooter.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置文本
    /// </summary>
    public string Text
    {
        get => _headerFooter?.Text ?? string.Empty;
        set
        {
            if (_headerFooter != null)
                _headerFooter.Text = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置格式
    /// </summary>
    public int Format
    {
        get => (int)(_headerFooter?.Format ?? 0);
        set
        {
            if (_headerFooter != null)
                _headerFooter.Format = (MsPowerPoint.PpDateTimeFormat)value;
        }
    }

    /// <summary>
    /// 获取或设置是否使用格式
    /// </summary>
    public bool UseFormat
    {
        get => _headerFooter?.UseFormat == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headerFooter != null)
                _headerFooter.UseFormat = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _headerFooter?.Parent;


    /// <summary>
    /// 获取或设置位置
    /// </summary>
    public int Position
    {
        get => 0; // HeaderFooter 没有直接的位置属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置对齐方式
    /// </summary>
    public int Alignment
    {
        get => 0; // HeaderFooter 没有直接的对齐方式属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置字体
    /// </summary>
    public IPowerPointFont Font
    {
        get
        {
            if (_font == null && _headerFooter?.Parent is MsPowerPoint.TextRange textRange)
            {
                _font = new PowerPointFont(textRange.Font);
            }
            return _font;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="headerFooter">COM HeaderFooter 对象</param>
    internal PowerPointHeaderFooter(MsPowerPoint.HeaderFooter headerFooter)
    {
        _headerFooter = headerFooter;
        _disposedValue = false;
    }

    /// <summary>
    /// 更新页眉页脚
    /// </summary>
    public void Update()
    {
        try
        {
            // HeaderFooter 通常是自动更新的
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to update header or footer.", ex);
        }
    }

    /// <summary>
    /// 应用到指定幻灯片
    /// </summary>
    /// <param name="slide">目标幻灯片</param>
    public void ApplyTo(IPowerPointSlide slide)
    {
        if (slide == null)
            throw new ArgumentNullException(nameof(slide));

        try
        {
            // 这需要具体的实现来应用页眉页脚到幻灯片
            throw new NotImplementedException("Applying header or footer to slide is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply header or footer to slide.", ex);
        }
    }

    /// <summary>
    /// 设置文本和格式
    /// </summary>
    /// <param name="text">文本内容</param>
    /// <param name="format">格式</param>
    /// <param name="useFormat">是否使用格式</param>
    public void SetTextAndFormat(string text, int format = 0, bool useFormat = false)
    {
        try
        {
            if (_headerFooter != null)
            {
                _headerFooter.Text = text ?? string.Empty;
                _headerFooter.Format = (PpDateTimeFormat)format;
                _headerFooter.UseFormat = useFormat ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set text and format for header or footer.", ex);
        }
    }

    /// <summary>
    /// 获取页眉页脚项信息
    /// </summary>
    /// <returns>页眉页脚项信息字符串</returns>
    public string GetHeaderFooterInfo()
    {
        try
        {
            return $"HeaderFooter - Text: {Text}, Visible: {Visible}, Format: {Format}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get header or footer info.", ex);
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
            _font?.Dispose();
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

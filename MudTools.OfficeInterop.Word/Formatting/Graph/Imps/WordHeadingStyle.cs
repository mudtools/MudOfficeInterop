
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordHeadingStyle : IWordHeadingStyle
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordHeadingStyle));
    private MsWord.HeadingStyle? _headingStyle;
    private bool _disposedValue;

    internal WordHeadingStyle(MsWord.HeadingStyle headingStyle)
    {
        _headingStyle = headingStyle ?? throw new ArgumentNullException(nameof(headingStyle));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _headingStyle != null ? new WordApplication(_headingStyle.Application) : null;

    public object? Parent => _headingStyle?.Parent;

    public int Level
    {
        get => _headingStyle?.Level ?? 0;
        set
        {
            if (_headingStyle != null)
                _headingStyle.Level = (short)value; // COM API uses short
        }
    }

    public IWordStyle? Style
    {
        get
        {
            if (_headingStyle == null)
                return null;
            if (_headingStyle.get_Style() is MsWord.Style style)
                return new WordStyle(style);
            return null;
        }
        set
        {
            if (_headingStyle != null)
            {
                var comStyle = (value as WordStyle)?.InternalComObject ?? null;
                _headingStyle.set_Style(comStyle);
            }
        }
    }
    #endregion

    public void Delete()
    {
        try
        {
            _headingStyle?.Delete();
        }
        catch (Exception ex)
        {
            log.Error("删除标题样式失败。", ex);
            throw new InvalidOperationException("删除标题样式失败。", ex);
        }
    }

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing && _headingStyle != null)
        {
            Marshal.ReleaseComObject(_headingStyle);
            _headingStyle = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
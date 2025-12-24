
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordHeadingStyles : IWordHeadingStyles
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordHeadingStyles));
    private readonly DisposableList _disposableList = new();
    private MsWord.HeadingStyles? _headingStyles;
    private bool _disposedValue;

    internal WordHeadingStyles(MsWord.HeadingStyles headingStyles)
    {
        _headingStyles = headingStyles ?? throw new ArgumentNullException(nameof(headingStyles));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _headingStyles != null ? new WordApplication(_headingStyles.Application) : null;

    public object? Parent => _headingStyles?.Parent;

    public int Count => _headingStyles?.Count ?? 0;

    public IWordHeadingStyle? this[int index]
    {
        get
        {
            if (_headingStyles == null || index < 1 || index > Count) return null;
            try
            {
                var comHeadingStyle = _headingStyles[index];
                var result = comHeadingStyle != null ? new WordHeadingStyle(comHeadingStyle) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索 HeadingStyle 对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    public IWordHeadingStyle Add(IWordStyle style, int level)
    {
        if (style == null)
            throw new ArgumentNullException(nameof(style));

        if (level < 1 || level > 9)
            throw new ArgumentOutOfRangeException(nameof(level), "目录级别必须在 1 到 9 之间。");

        if (_headingStyles == null)
            throw new InvalidOperationException("标题样式集合不可用。");

        try
        {
            var comStyle = (style as WordStyle)?.InternalComObject ?? null;

            var comHeadingStyle = _headingStyles.Add(comStyle, (short)level);
            var wrapper = new WordHeadingStyle(comHeadingStyle);
            return wrapper;
        }
        catch (Exception ex)
        {
            log.Error("向 HeadingStyles 集合添加新样式失败。", ex);
            throw new InvalidOperationException("添加标题样式失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _headingStyles != null)
        {
            Marshal.ReleaseComObject(_headingStyles);
            _disposableList.Dispose();
            _headingStyles = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordHeadingStyle> 实现

    public IEnumerator<IWordHeadingStyle> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
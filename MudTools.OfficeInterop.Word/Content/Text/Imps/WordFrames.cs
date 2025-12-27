//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Frames 的封装实现类。
/// </summary>
internal class WordFrames : IWordFrames
{
    private MsWord.Frames _frames;
    private bool _disposedValue;

    internal WordFrames(MsWord.Frames frames)
    {
        _frames = frames ?? throw new ArgumentNullException(nameof(frames));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _frames != null ? new WordApplication(_frames.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _frames?.Parent;

    /// <inheritdoc/>
    public int Count => _frames?.Count ?? 0;

    /// <inheritdoc/>
    public IWordFrame First => _frames?.Count > 0 ? new WordFrame(_frames[1]) : null;

    /// <inheritdoc/>
    public IWordFrame Last => _frames?.Count > 0 ? new WordFrame(_frames[_frames.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordFrame this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comFrame = _frames[index];
                return new WordFrame(comFrame);
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordFrame Add(IWordRange range)
    {
        if (_frames == null || range == null) return null;

        try
        {
            var comRange = (range as WordRange)?.InternalComObject;
            if (comRange != null)
            {
                var newFrame = _frames.Add(comRange);
                return new WordFrame(newFrame);
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加框架。", ex);
        }

        return null;
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        _frames[index].Delete();
    }

    /// <inheritdoc/>
    public void DeleteRange(int startIndex, int count)
    {
        if (startIndex < 1 || count <= 0) return;

        int endIndex = Math.Min(startIndex + count - 1, Count);
        if (startIndex <= endIndex)
        {
            // 从后往前删除，避免索引变化
            for (int i = endIndex; i >= startIndex; i--)
            {
                _frames[i].Delete();
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_frames == null) return;

        // 从后往前删除，避免索引变化
        for (int i = Count; i >= 1; i--)
        {
            _frames[i].Delete();
        }
    }

    /// <inheritdoc/>
    public List<int> GetIndexes()
    {
        var indexes = new List<int>();
        for (int i = 1; i <= Count; i++)
        {
            indexes.Add(i);
        }
        return indexes;
    }

    /// <inheritdoc/>
    public List<IWordFrame> GetRange(int startIndex, int endIndex)
    {
        var frames = new List<IWordFrame>();
        if (_frames == null) return frames;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        try
        {
            for (int i = validStart; i <= validEnd; i++)
            {
                frames.Add(new WordFrame(_frames[i]));
            }
        }
        catch
        {
            // 获取范围失败返回已获取的框架
        }

        return frames;
    }

    /// <inheritdoc/>
    public void Select()
    {
        // 选择所有框架
        if (Count > 0)
        {
            for (int i = 1; i <= Count; i++)
            {
                _frames[i].Select();
            }
        }
    }

    /// <inheritdoc/>
    public void Copy()
    {
        try
        {
            _frames.Application.Selection.Copy();
        }
        catch
        {
            // 复制失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Cut()
    {
        try
        {
            _frames.Application.Selection.Cut();
        }
        catch
        {
            // 剪切失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Paste()
    {
        try
        {
            _frames.Application.Selection.Paste();
        }
        catch
        {
            // 粘贴失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_frames == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var frame = _frames[i];
                if (frame?.Borders != null)
                {
                    frame.Borders.Enable = 1;
                    foreach (MsWord.Border border in frame.Borders)
                    {
                        border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                        border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                        border.Color = (MsWord.WdColor)(int)color;
                    }
                }
            }
        }
        catch
        {
            // 设置边框失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void SetShadingForAll(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_frames == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var frame = _frames[i];
                if (frame?.Shading != null)
                {
                    frame.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                    if (foregroundColor != WdColor.wdColorAutomatic)
                        frame.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                    if (backgroundColor != WdColor.wdColorWhite)
                        frame.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;

                }
            }
        }
        catch
        {
            // 设置底纹失败忽略异常
        }
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordFrame> GetEnumerator()
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

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _frames != null)
        {
            Marshal.ReleaseComObject(_frames);
            _frames = null;
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
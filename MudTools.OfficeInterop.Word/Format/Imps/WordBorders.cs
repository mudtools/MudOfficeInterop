//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Borders 的实现类。
/// </summary>
internal class WordBorders : IWordBorders
{
    private MsWord.Borders _borders;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="borders">原始 COM Borders 对象。</param>
    internal WordBorders(MsWord.Borders borders)
    {
        _borders = borders ?? throw new ArgumentNullException(nameof(borders));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _borders != null ? new WordApplication(_borders.Application) : null;

    /// <inheritdoc/>
    public object Parent => _borders?.Parent;

    /// <inheritdoc/>
    public int Count => _borders?.Count ?? 0;

    /// <inheritdoc/>
    public IWordBorder this[WdBorderType borderType]
    {
        get
        {
            if (_borders == null)
                return null;

            try
            {
                var border = _borders[(MsWord.WdBorderType)(int)borderType];
                return border != null ? new WordBorder(border) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public bool Enable
    {
        get => _borders?.Enable == 1;
        set
        {
            if (_borders != null)
                _borders.Enable = value ? 1 : 0;
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ApplyStyle(WdLineStyle lineStyle, WdLineWidth lineWidth, WdColor color)
    {
        if (_borders == null)
            throw new ObjectDisposedException(nameof(WordBorders));

        try
        {
            foreach (MsWord.Border border in _borders)
            {
                if (border != null)
                {
                    border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                    border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                    border.Color = (MsWord.WdColor)(int)color;
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法应用边框样式。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Contains(WdBorderType borderType)
    {
        if (_borders == null)
            return false;

        return _borders[(MsWord.WdBorderType)(int)borderType] != null;
    }

    /// <inheritdoc/>
    public List<WdBorderType> GetBorderTypes()
    {
        var types = new List<WdBorderType>();

        if (_borders == null)
            return types;

        // 常见的边框类型枚举
        var allTypes = new[]
        {
            WdBorderType.wdBorderTop,
            WdBorderType.wdBorderLeft,
            WdBorderType.wdBorderBottom,
            WdBorderType.wdBorderRight,
            WdBorderType.wdBorderHorizontal,
            WdBorderType.wdBorderVertical,
            WdBorderType.wdBorderDiagonalDown,
            WdBorderType.wdBorderDiagonalUp
        };

        foreach (var type in allTypes)
        {
            if (Contains(type))
                types.Add(type);
        }

        return types;
    }

    #endregion

    #region IEnumerable<IWordBorder> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordBorder> GetEnumerator()
    {
        if (_borders == null)
            yield break;

        foreach (var border in _borders)
        {
            var b = border as MsWord.Border;
            if (border != null)
                yield return new WordBorder(b);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _borders != null)
        {
            Marshal.ReleaseComObject(_borders);
            _borders = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Word 文档中一组边框（Borders）的封装实现类。
/// </summary>
internal class WordBorders : IWordBorders
{
    internal MsWord.Borders _borders;

    internal MsWord.Borders InternalComObject => _borders;
    private bool _disposedValue;
    private DisposableList _disposables = [];

    /// <summary>
    /// 初始化 <see cref="WordBorders"/> 类的新实例。
    /// </summary>
    /// <param name="borders">要封装的原始 COM Borders 对象。</param>
    internal WordBorders(MsWord.Borders borders)
    {
        _borders = borders ?? throw new ArgumentNullException(nameof(borders));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _borders != null ? new WordApplication(_borders.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _borders?.Parent;

    /// <inheritdoc/>
    public int Count => _borders?.Count ?? 0;

    /// <inheritdoc/>
    public IWordBorder? this[WdBorderType index]
    {
        get
        {
            if (_borders == null) return null;
            try
            {
                var comBorder = _borders[(MsWord.WdBorderType)(int)index];
                var border = comBorder != null ? new WordBorder(comBorder) : null;
                if (border != null)
                    _disposables.Add(border);
                return border;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public bool Enable
    {
        get => _borders?.Enable != null && _borders?.Enable == 1;
        set { if (_borders != null) _borders.Enable = value ? 1 : 0; }
    }

    public WdLineStyle LineStyle
    {
        get => _borders?.OutsideLineStyle != null ? _borders.OutsideLineStyle.EnumConvert(WdLineStyle.wdLineStyleSingle) : WdLineStyle.wdLineStyleSingle;
        set
        {
            if (_borders != null) _borders.OutsideLineStyle = value.EnumConvert(MsWord.WdLineStyle.wdLineStyleSingle);
        }
    }

    public WdLineWidth LineWidth
    {
        get => _borders?.OutsideLineWidth != null ? _borders.OutsideLineWidth.EnumConvert(WdLineWidth.wdLineWidth100pt) : WdLineWidth.wdLineWidth100pt;
        set
        {
            if (_borders != null) _borders.OutsideLineWidth = value.EnumConvert(MsWord.WdLineWidth.wdLineWidth100pt);
        }
    }

    /// <inheritdoc/>
    public bool JoinBorders
    {
        get => _borders?.JoinBorders ?? false;
        set { if (_borders != null) _borders.JoinBorders = value; }
    }

    /// <inheritdoc/>
    public WdColor InsideColor
    {
        get => _borders?.InsideColor != null ? _borders.InsideColor.EnumConvert(WdColor.wdColorAutomatic) : WdColor.wdColorAutomatic;
        set
        {
            if (_borders != null) _borders.InsideColor = value.EnumConvert(MsWord.WdColor.wdColorAutomatic);
        }
    }

    /// <inheritdoc/>
    public WdColorIndex InsideColorIndex
    {
        get => _borders?.InsideColorIndex != null ? _borders.InsideColorIndex.EnumConvert(WdColorIndex.wdAuto) : WdColorIndex.wdAuto;
        set
        {
            if (_borders != null) _borders.InsideColorIndex = value.EnumConvert(MsWord.WdColorIndex.wdAuto);
        }
    }

    /// <inheritdoc/>
    public WdLineStyle InsideLineStyle
    {
        get => _borders?.InsideLineStyle != null ? _borders.InsideLineStyle.EnumConvert(WdLineStyle.wdLineStyleSingle) : WdLineStyle.wdLineStyleSingle;
        set
        {
            if (_borders != null) _borders.InsideLineStyle = value.EnumConvert(MsWord.WdLineStyle.wdLineStyleSingle);
        }
    }

    /// <inheritdoc/>
    public WdLineWidth InsideLineWidth
    {
        get => _borders?.OutsideLineWidth != null ? _borders.OutsideLineWidth.EnumConvert(WdLineWidth.wdLineWidth100pt) : WdLineWidth.wdLineWidth100pt;
        set
        {
            if (_borders != null) _borders.OutsideLineWidth = value.EnumConvert(MsWord.WdLineWidth.wdLineWidth100pt);
        }
    }

    /// <inheritdoc/>
    public WdColor OutsideColor
    {
        get => _borders?.OutsideColor != null ? _borders.OutsideColor.EnumConvert(WdColor.wdColorAutomatic) : WdColor.wdColorAutomatic;
        set
        {
            if (_borders != null) _borders.OutsideColor = value.EnumConvert(MsWord.WdColor.wdColorAutomatic);
        }
    }

    /// <inheritdoc/>
    public WdColorIndex OutsideColorIndex
    {
        get => _borders?.OutsideColorIndex != null ? _borders.OutsideColorIndex.EnumConvert(WdColorIndex.wdAuto) : WdColorIndex.wdAuto;
        set
        {
            if (_borders != null) _borders.OutsideColorIndex = value.EnumConvert(MsWord.WdColorIndex.wdAuto);
        }
    }

    /// <inheritdoc/>
    public WdLineStyle OutsideLineStyle
    {
        get => _borders?.OutsideLineStyle != null ? _borders.OutsideLineStyle.EnumConvert(WdLineStyle.wdLineStyleSingle) : WdLineStyle.wdLineStyleSingle;
        set
        {
            if (_borders != null) _borders.OutsideLineStyle = value.EnumConvert(MsWord.WdLineStyle.wdLineStyleSingle);
        }
    }

    /// <inheritdoc/>
    public WdLineWidth OutsideLineWidth
    {
        get => _borders?.OutsideLineWidth != null ? _borders.OutsideLineWidth.EnumConvert(WdLineWidth.wdLineWidth100pt) : WdLineWidth.wdLineWidth100pt;
        set
        {
            if (_borders != null) _borders.OutsideLineWidth = value.EnumConvert(MsWord.WdLineWidth.wdLineWidth100pt);
        }
    }

    /// <inheritdoc/>
    public bool HasHorizontal => _borders?.HasHorizontal ?? false;

    /// <inheritdoc/>
    public bool HasVertical => _borders?.HasVertical ?? false;

    /// <inheritdoc/>
    public bool AlwaysInFront
    {
        get => _borders?.AlwaysInFront ?? false;
        set { if (_borders != null) _borders.AlwaysInFront = value; }
    }

    /// <inheritdoc/>
    public WdBorderDistanceFrom DistanceFrom
    {
        get => _borders?.DistanceFrom != null ? _borders.DistanceFrom.EnumConvert(WdBorderDistanceFrom.wdBorderDistanceFromText) : WdBorderDistanceFrom.wdBorderDistanceFromText;
        set
        {
            if (_borders != null) _borders.DistanceFrom = value.EnumConvert(MsWord.WdBorderDistanceFrom.wdBorderDistanceFromText);
        }
    }

    /// <inheritdoc/>
    public int DistanceFromBottom
    {
        get => _borders?.DistanceFromBottom ?? 0;
        set { if (_borders != null) _borders.DistanceFromBottom = value; }
    }

    /// <inheritdoc/>
    public int DistanceFromLeft
    {
        get => _borders?.DistanceFromLeft ?? 0;
        set { if (_borders != null) _borders.DistanceFromLeft = value; }
    }

    /// <inheritdoc/>
    public int DistanceFromRight
    {
        get => _borders?.DistanceFromRight ?? 0;
        set { if (_borders != null) _borders.DistanceFromRight = value; }
    }

    /// <inheritdoc/>
    public int DistanceFromTop
    {
        get => _borders?.DistanceFromTop ?? 0;
        set { if (_borders != null) _borders.DistanceFromTop = value; }
    }

    /// <inheritdoc/>
    public bool EnableFirstPageInSection
    {
        get => _borders?.EnableFirstPageInSection ?? false;
        set { if (_borders != null) _borders.EnableFirstPageInSection = value; }
    }

    /// <inheritdoc/>
    public bool EnableOtherPagesInSection
    {
        get => _borders?.EnableOtherPagesInSection ?? false;
        set { if (_borders != null) _borders.EnableOtherPagesInSection = value; }
    }

    /// <inheritdoc/>
    public bool SurroundFooter
    {
        get => _borders?.SurroundFooter ?? false;
        set { if (_borders != null) _borders.SurroundFooter = value; }
    }

    /// <inheritdoc/>
    public bool SurroundHeader
    {
        get => _borders?.SurroundHeader ?? false;
        set { if (_borders != null) _borders.SurroundHeader = value; }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ApplyPageBordersToAllSections()
    {
        _borders?.ApplyPageBordersToAllSections();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordBorders"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _borders != null)
        {
            _disposables.Dispose();
            Marshal.ReleaseComObject(_borders);
            _borders = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordBorders"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordBorder> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordBorder> GetEnumerator()
    {
        foreach (var b in _borders)
        {
            yield return new WordBorder(b as MsWord.Border);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
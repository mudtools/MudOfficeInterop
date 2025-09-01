//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office线条格式对象的实现类
/// </summary>
internal class OfficeLineFormat : IOfficeLineFormat
{
    private MsCore.LineFormat _lineFormat;
    private bool _disposedValue;

    /// <summary>
    /// 初始化OfficeLineFormat类的新实例
    /// </summary>
    /// <param name="lineFormat">原始的COM线条格式对象</param>
    internal OfficeLineFormat(MsCore.LineFormat lineFormat)
    {
        _lineFormat = lineFormat ?? throw new ArgumentNullException(nameof(lineFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public bool Visible
    {
        get => _lineFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public float Weight
    {
        get => _lineFormat?.Weight ?? 0;
        set
        {
            if (_lineFormat != null && value > 0)
                _lineFormat.Weight = value;
        }
    }

    /// <inheritdoc/>
    public MsoLineDashStyle DashStyle
    {
        get => _lineFormat?.DashStyle != null ? (MsoLineDashStyle)(int)_lineFormat?.DashStyle : MsoLineDashStyle.msoLineDashStyleMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.DashStyle = (MsCore.MsoLineDashStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoLineStyle Style
    {
        get => _lineFormat?.Style != null ? (MsoLineStyle)(int)_lineFormat?.Style : MsoLineStyle.msoLineStyleMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.Style = (MsCore.MsoLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public IOfficeColorFormat ForeColor
    {
        get
        {
            if (_lineFormat == null)
                return null;

            var foreColor = _lineFormat.ForeColor;
            return foreColor != null ? new OfficeColorFormat(foreColor) : null;
        }
    }

    /// <inheritdoc/>
    public IOfficeColorFormat BackColor
    {
        get
        {
            if (_lineFormat == null)
                return null;

            var backColor = _lineFormat.BackColor;
            return backColor != null ? new OfficeColorFormat(backColor) : null;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _lineFormat?.Transparency ?? 0;
        set
        {
            if (_lineFormat != null && value >= 0 && value <= 1)
                _lineFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadLength BeginArrowheadLength
    {
        get => _lineFormat?.BeginArrowheadLength != null ? (MsoArrowheadLength)(int)_lineFormat?.BeginArrowheadLength : MsoArrowheadLength.msoArrowheadLengthMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.BeginArrowheadLength = (MsCore.MsoArrowheadLength)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadStyle BeginArrowheadStyle
    {
        get => _lineFormat?.BeginArrowheadStyle != null ? (MsoArrowheadStyle)(int)_lineFormat?.BeginArrowheadStyle : MsoArrowheadStyle.msoArrowheadNone;
        set
        {
            if (_lineFormat != null) _lineFormat.BeginArrowheadStyle = (MsCore.MsoArrowheadStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadWidth BeginArrowheadWidth
    {
        get => _lineFormat?.BeginArrowheadWidth != null ? (MsoArrowheadWidth)(int)_lineFormat?.BeginArrowheadWidth : MsoArrowheadWidth.msoArrowheadWidthMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.BeginArrowheadWidth = (MsCore.MsoArrowheadWidth)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadLength EndArrowheadLength
    {
        get => _lineFormat?.EndArrowheadLength != null ? (MsoArrowheadLength)(int)_lineFormat?.EndArrowheadLength : MsoArrowheadLength.msoArrowheadLengthMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.EndArrowheadLength = (MsCore.MsoArrowheadLength)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadStyle EndArrowheadStyle
    {
        get => _lineFormat?.EndArrowheadStyle != null ? (MsoArrowheadStyle)(int)_lineFormat?.EndArrowheadStyle : MsoArrowheadStyle.msoArrowheadStyleMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.EndArrowheadStyle = (MsCore.MsoArrowheadStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoArrowheadWidth EndArrowheadWidth
    {
        get => _lineFormat?.EndArrowheadWidth != null ? (MsoArrowheadWidth)(int)_lineFormat?.EndArrowheadWidth : MsoArrowheadWidth.msoArrowheadNarrow;
        set
        {
            if (_lineFormat != null) _lineFormat.EndArrowheadWidth = (MsCore.MsoArrowheadWidth)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoPatternType Pattern
    {
        get => _lineFormat?.Pattern != null ? (MsoPatternType)(int)_lineFormat?.Pattern : MsoPatternType.msoPatternMixed;
        set
        {
            if (_lineFormat != null) _lineFormat.Pattern = (MsCore.MsoPatternType)(int)value;
        }
    }
    #endregion

    #region IDisposable实现

    /// <summary>
    /// 释放资源:cite[6]
    /// </summary>
    /// <param name="disposing">是否正在处置</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _lineFormat != null)
        {
            // 释放COM对象资源:cite[1]
            Marshal.ReleaseComObject(_lineFormat);
            _lineFormat = null;
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
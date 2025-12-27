//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Core.LineFormat 的实现类。
/// </summary>
internal class WordLineFormat : IWordLineFormat
{
    private MsWord.LineFormat _lineFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="lineFormat">原始 COM LineFormat 对象。</param>
    internal WordLineFormat(MsWord.LineFormat lineFormat)
    {
        _lineFormat = lineFormat ?? throw new ArgumentNullException(nameof(lineFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _lineFormat != null ? new WordApplication(_lineFormat.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _lineFormat?.Parent;

    /// <inheritdoc/>
    public IWordColorFormat ForeColor
    {
        get => _lineFormat?.ForeColor != null ? new WordColorFormat(_lineFormat.ForeColor) : null;
    }

    /// <inheritdoc/>
    public IWordColorFormat BackColor
    {
        get => _lineFormat?.BackColor != null ? new WordColorFormat(_lineFormat.BackColor) : null;

    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _lineFormat?.Transparency ?? 0f;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Transparency = value;
        }
    }

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
        get => _lineFormat?.Weight ?? 0f;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Weight = value;
        }
    }

    /// <inheritdoc/>
    public MsoLineStyle Style
    {
        get => _lineFormat?.Style != null ? (MsoLineStyle)(int)_lineFormat?.Style : MsoLineStyle.msoLineSingle;
        set
        {
            if (_lineFormat != null) _lineFormat.Style = (MsCore.MsoLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoLineDashStyle DashStyle
    {
        get => _lineFormat?.DashStyle != null ? (MsoLineDashStyle)(int)_lineFormat?.DashStyle : MsoLineDashStyle.msoLineSolid;
        set => _lineFormat.DashStyle = (MsCore.MsoLineDashStyle)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadStyle BeginArrowheadStyle
    {
        get => _lineFormat?.BeginArrowheadStyle != null ? (MsoArrowheadStyle)(int)_lineFormat?.BeginArrowheadStyle : MsoArrowheadStyle.msoArrowheadNone;
        set => _lineFormat.BeginArrowheadStyle = (MsCore.MsoArrowheadStyle)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadWidth BeginArrowheadWidth
    {
        get => _lineFormat?.BeginArrowheadWidth != null ? (MsoArrowheadWidth)(int)_lineFormat?.BeginArrowheadWidth : MsoArrowheadWidth.msoArrowheadWidthMixed;
        set => _lineFormat.BeginArrowheadWidth = (MsCore.MsoArrowheadWidth)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadLength BeginArrowheadLength
    {
        get => _lineFormat?.BeginArrowheadLength != null ? (MsoArrowheadLength)(int)_lineFormat?.BeginArrowheadLength : MsoArrowheadLength.msoArrowheadLengthMixed;
        set => _lineFormat.BeginArrowheadLength = (MsCore.MsoArrowheadLength)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadStyle EndArrowheadStyle
    {
        get => _lineFormat?.EndArrowheadStyle != null ? (MsoArrowheadStyle)(int)_lineFormat?.EndArrowheadStyle : MsoArrowheadStyle.msoArrowheadStyleMixed;
        set => _lineFormat.EndArrowheadStyle = (MsCore.MsoArrowheadStyle)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadWidth EndArrowheadWidth
    {
        get => _lineFormat?.EndArrowheadWidth != null ? (MsoArrowheadWidth)(int)_lineFormat?.EndArrowheadWidth : MsoArrowheadWidth.msoArrowheadWidthMixed;
        set => _lineFormat.EndArrowheadWidth = (MsCore.MsoArrowheadWidth)(int)value;
    }

    /// <inheritdoc/>
    public MsoArrowheadLength EndArrowheadLength
    {
        get => _lineFormat?.EndArrowheadLength != null ? (MsoArrowheadLength)(int)_lineFormat?.EndArrowheadLength : MsoArrowheadLength.msoArrowheadLengthMixed;
        set => _lineFormat.EndArrowheadLength = (MsCore.MsoArrowheadLength)(int)value;
    }

    /// <inheritdoc/>
    public MsoPatternType Pattern
    {
        get => _lineFormat?.Pattern != null ? (MsoPatternType)(int)_lineFormat?.Pattern : MsoPatternType.msoPattern10Percent;
        set => _lineFormat.Pattern = (MsCore.MsoPatternType)(int)value;
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Solid(int color)
    {
        if (_lineFormat != null)
        {
            _lineFormat.Style = MsCore.MsoLineStyle.msoLineSingle;
            if (_lineFormat.ForeColor != null)
                _lineFormat.ForeColor.RGB = color;
        }
    }

    /// <inheritdoc/>
    public void SetDashStyle(MsoLineDashStyle dashStyle)
    {
        if (_lineFormat != null)
        {
            _lineFormat.DashStyle = (MsCore.MsoLineDashStyle)(int)dashStyle;
        }
    }


    /// <inheritdoc/>
    public void Clear()
    {
        if (_lineFormat != null)
        {
            _lineFormat.Visible = MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordLineFormat targetLine)
    {
        if (_lineFormat == null || targetLine == null)
            return;

        try
        {
            targetLine.Visible = this.Visible;
            targetLine.Weight = this.Weight;
            targetLine.Style = this.Style;
            targetLine.DashStyle = this.DashStyle;
            targetLine.Transparency = this.Transparency;

            targetLine.BeginArrowheadStyle = this.BeginArrowheadStyle;
            targetLine.BeginArrowheadWidth = this.BeginArrowheadWidth;
            targetLine.BeginArrowheadLength = this.BeginArrowheadLength;
            targetLine.EndArrowheadStyle = this.EndArrowheadStyle;
            targetLine.EndArrowheadWidth = this.EndArrowheadWidth;
            targetLine.EndArrowheadLength = this.EndArrowheadLength;

            // 复制颜色
            if (this.ForeColor != null && targetLine.ForeColor != null)
            {
                targetLine.ForeColor.RGB = this.ForeColor.RGB;
            }
            if (this.BackColor != null && targetLine.BackColor != null)
            {
                targetLine.BackColor.RGB = this.BackColor.RGB;
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制线条格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_lineFormat != null)
        {
            _lineFormat.Visible = MsCore.MsoTriState.msoFalse;
            _lineFormat.Weight = 1.0f;
            _lineFormat.Style = MsCore.MsoLineStyle.msoLineSingle;
            _lineFormat.DashStyle = MsCore.MsoLineDashStyle.msoLineSolid;
            _lineFormat.Transparency = 0f;

            _lineFormat.BeginArrowheadStyle = MsCore.MsoArrowheadStyle.msoArrowheadNone;
            _lineFormat.BeginArrowheadWidth = MsCore.MsoArrowheadWidth.msoArrowheadWidthMedium;
            _lineFormat.BeginArrowheadLength = MsCore.MsoArrowheadLength.msoArrowheadLengthMedium;
            _lineFormat.EndArrowheadStyle = MsCore.MsoArrowheadStyle.msoArrowheadNone;
            _lineFormat.EndArrowheadWidth = MsCore.MsoArrowheadWidth.msoArrowheadWidthMedium;
            _lineFormat.EndArrowheadLength = MsCore.MsoArrowheadLength.msoArrowheadLengthMedium;
        }
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

        if (disposing)
        {
            // 释放前景色对象
            if (_lineFormat?.ForeColor != null)
            {
                Marshal.ReleaseComObject(_lineFormat.ForeColor);
            }
            // 释放背景色对象
            if (_lineFormat?.BackColor != null)
            {
                Marshal.ReleaseComObject(_lineFormat.BackColor);
            }
            // 释放线条格式对象本身
            if (_lineFormat != null)
            {
                Marshal.ReleaseComObject(_lineFormat);
                _lineFormat = null;
            }
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
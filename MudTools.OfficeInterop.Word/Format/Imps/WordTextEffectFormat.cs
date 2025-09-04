//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表示与艺术字对象关联的文本效果格式的封装实现类。
/// </summary>
internal class WordTextEffectFormat : IWordTextEffectFormat
{
    private MsWord.TextEffectFormat _textEffectFormat;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordTextEffectFormat"/> 类的新实例。
    /// </summary>
    /// <param name="textEffectFormat">要封装的原始 COM TextEffectFormat 对象。</param>
    internal WordTextEffectFormat(MsWord.TextEffectFormat textEffectFormat)
    {
        _textEffectFormat = textEffectFormat ?? throw new ArgumentNullException(nameof(textEffectFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _textEffectFormat != null ? new WordApplication(_textEffectFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _textEffectFormat?.Parent;

    /// <inheritdoc/>
    public MsoTextEffectAlignment Alignment
    {
        get => _textEffectFormat?.Alignment != null ? (MsoTextEffectAlignment)(int)_textEffectFormat?.Alignment : MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
        set
        {
            if (_textEffectFormat != null) _textEffectFormat.Alignment = (MsCore.MsoTextEffectAlignment)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool FontBold
    {
        get => _textEffectFormat?.FontBold != null && _textEffectFormat?.FontBold == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontBold = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool FontItalic
    {
        get => _textEffectFormat?.FontItalic != null && _textEffectFormat?.FontItalic == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontItalic = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public string FontName
    {
        get => _textEffectFormat?.FontName ?? string.Empty;
        set { if (_textEffectFormat != null) _textEffectFormat.FontName = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public float FontSize
    {
        get => _textEffectFormat?.FontSize ?? 0f;
        set { if (_textEffectFormat != null) _textEffectFormat.FontSize = value; }
    }

    /// <inheritdoc/>
    public bool KernedPairs
    {
        get => _textEffectFormat?.KernedPairs != null && _textEffectFormat?.KernedPairs == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.KernedPairs = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool NormalizedHeight
    {
        get => _textEffectFormat?.NormalizedHeight != null && _textEffectFormat?.NormalizedHeight == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.NormalizedHeight = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public MsoPresetTextEffectShape PresetShape
    {
        get => _textEffectFormat?.PresetShape != null ? (MsoPresetTextEffectShape)(int)_textEffectFormat?.PresetShape : MsoPresetTextEffectShape.msoTextEffectShapeMixed;
        set
        {
            if (_textEffectFormat != null) _textEffectFormat.PresetShape = (MsCore.MsoPresetTextEffectShape)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoPresetTextEffect PresetTextEffect
    {
        get => _textEffectFormat?.PresetTextEffect != null ? (MsoPresetTextEffect)(int)_textEffectFormat?.PresetTextEffect : MsoPresetTextEffect.msoTextEffectMixed;
        set
        {
            if (_textEffectFormat != null) _textEffectFormat.PresetTextEffect = (MsCore.MsoPresetTextEffect)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool RotatedChars
    {
        get => _textEffectFormat?.RotatedChars != null && _textEffectFormat?.RotatedChars == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.RotatedChars = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public string Text
    {
        get => _textEffectFormat?.Text ?? string.Empty;
        set { if (_textEffectFormat != null) _textEffectFormat.Text = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public float Tracking
    {
        get => _textEffectFormat?.Tracking ?? 0f;
        set { if (_textEffectFormat != null) _textEffectFormat.Tracking = value; }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ToggleVerticalText()
    {
        _textEffectFormat?.ToggleVerticalText();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTextEffectFormat"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _textEffectFormat != null)
        {
            Marshal.ReleaseComObject(_textEffectFormat);
            _textEffectFormat = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordTextEffectFormat"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
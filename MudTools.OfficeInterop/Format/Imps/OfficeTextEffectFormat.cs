//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.TextEffectFormat 的二次封装实现类。
/// 提供安全访问艺术字格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeTextEffectFormat : IOfficeTextEffectFormat
{
    private MsCore.TextEffectFormat _textEffectFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 TextEffectFormat 对象。
    /// </summary>
    /// <param name="textEffectFormat">原始的 COM TextEffectFormat 对象。</param>
    internal OfficeTextEffectFormat(MsCore.TextEffectFormat textEffectFormat)
    {
        _textEffectFormat = textEffectFormat ?? throw new ArgumentNullException(nameof(textEffectFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string Text
    {
        get => _textEffectFormat?.Text ?? string.Empty;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.Text = value;
        }
    }

    /// <inheritdoc/>
    public string FontName
    {
        get => _textEffectFormat?.FontName ?? string.Empty;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontName = value;
        }
    }

    /// <inheritdoc/>
    public float FontSize
    {
        get => _textEffectFormat?.FontSize ?? 0f;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontSize = value;
        }
    }

    /// <inheritdoc/>
    public bool FontBold
    {
        get => _textEffectFormat?.FontBold == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontBold = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool FontItalic
    {
        get => _textEffectFormat?.FontItalic == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontItalic = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

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
    public MsoPresetTextEffect PresetTextEffect
    {
        get => _textEffectFormat?.PresetTextEffect != null ? (MsoPresetTextEffect)(int)_textEffectFormat?.PresetTextEffect : MsoPresetTextEffect.msoTextEffectMixed;
        set
        {
            if (_textEffectFormat != null) _textEffectFormat.PresetTextEffect = (MsCore.MsoPresetTextEffect)(int)value;
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
    public bool RotatedChars
    {
        get => _textEffectFormat?.RotatedChars == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.RotatedChars = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool NormalizedHeight
    {
        get => _textEffectFormat?.NormalizedHeight == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.NormalizedHeight = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public float Tracking
    {
        get => _textEffectFormat?.Tracking ?? 0f;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.Tracking = value;
        }
    }

    #endregion

    #region 方法实现
    public void ToggleVerticalText()
    {
        _textEffectFormat?.ToggleVerticalText();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
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

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Microsoft Word 中自动更正功能的封装实现类。
/// </summary>
internal class WordAutoCorrect : IWordAutoCorrect
{
    private MsWord.AutoCorrect _autoCorrect;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordAutoCorrect"/> 类的新实例。
    /// </summary>
    /// <param name="autoCorrect">要封装的原始 COM AutoCorrect 对象。</param>
    internal WordAutoCorrect(MsWord.AutoCorrect autoCorrect)
    {
        _autoCorrect = autoCorrect ?? throw new ArgumentNullException(nameof(autoCorrect));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _autoCorrect != null ? new WordApplication(_autoCorrect.Application) : null;

    /// <inheritdoc/>
    public object Parent => _autoCorrect?.Parent;

    /// <inheritdoc/>
    public int Creator => _autoCorrect?.Creator ?? 0;

    #endregion

    #region 自动更正选项属性实现 (AutoCorrect Options Properties Implementation)

    /// <inheritdoc/>
    public bool CorrectCapsLock
    {
        get => _autoCorrect?.CorrectCapsLock ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectCapsLock = value; }
    }

    /// <inheritdoc/>
    public bool CorrectDays
    {
        get => _autoCorrect?.CorrectDays ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectDays = value; }
    }

    /// <inheritdoc/>
    public bool CorrectHangulAndAlphabet
    {
        get => _autoCorrect?.CorrectHangulAndAlphabet ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectHangulAndAlphabet = value; }
    }

    /// <inheritdoc/>
    public bool CorrectInitialCaps
    {
        get => _autoCorrect?.CorrectInitialCaps ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectInitialCaps = value; }
    }

    /// <inheritdoc/>
    public bool CorrectKeyboardSetting
    {
        get => _autoCorrect?.CorrectKeyboardSetting ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectKeyboardSetting = value; }
    }

    /// <inheritdoc/>
    public bool CorrectSentenceCaps
    {
        get => _autoCorrect?.CorrectSentenceCaps ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectSentenceCaps = value; }
    }

    /// <inheritdoc/>
    public bool CorrectTableCells
    {
        get => _autoCorrect?.CorrectTableCells ?? false;
        set { if (_autoCorrect != null) _autoCorrect.CorrectTableCells = value; }
    }

    /// <inheritdoc/>
    public bool DisplayAutoCorrectOptions
    {
        get => _autoCorrect?.DisplayAutoCorrectOptions ?? false;
        set { if (_autoCorrect != null) _autoCorrect.DisplayAutoCorrectOptions = value; }
    }

    /// <inheritdoc/>
    public IWordAutoCorrectEntries? Entries => _autoCorrect?.Entries != null ? new WordAutoCorrectEntries(_autoCorrect.Entries) : null;

    /// <inheritdoc/>
    public bool FirstLetterAutoAdd
    {
        get => _autoCorrect?.FirstLetterAutoAdd ?? false;
        set { if (_autoCorrect != null) _autoCorrect.FirstLetterAutoAdd = value; }
    }

    /// <inheritdoc/>
    public IWordFirstLetterExceptions? FirstLetterExceptions
        => _autoCorrect?.FirstLetterExceptions != null ? new WordFirstLetterExceptions(_autoCorrect.FirstLetterExceptions) : null;

    /// <inheritdoc/>
    public bool HangulAndAlphabetAutoAdd
    {
        get => _autoCorrect?.HangulAndAlphabetAutoAdd ?? false;
        set { if (_autoCorrect != null) _autoCorrect.HangulAndAlphabetAutoAdd = value; }
    }

    /// <inheritdoc/>
    public IWordHangulAndAlphabetExceptions? HangulAndAlphabetExceptions
        => _autoCorrect?.HangulAndAlphabetExceptions != null ? new WordHangulAndAlphabetExceptions(_autoCorrect.HangulAndAlphabetExceptions) : null;

    /// <inheritdoc/>
    public bool OtherCorrectionsAutoAdd
    {
        get => _autoCorrect?.OtherCorrectionsAutoAdd ?? false;
        set { if (_autoCorrect != null) _autoCorrect.OtherCorrectionsAutoAdd = value; }
    }

    /// <inheritdoc/>
    public IWordOtherCorrectionsExceptions? OtherCorrectionsExceptions
        => _autoCorrect?.OtherCorrectionsExceptions != null ? new WordOtherCorrectionsExceptions(_autoCorrect.OtherCorrectionsExceptions) : null;

    /// <inheritdoc/>
    public bool ReplaceText
    {
        get => _autoCorrect?.ReplaceText ?? false;
        set { if (_autoCorrect != null) _autoCorrect.ReplaceText = value; }
    }

    /// <inheritdoc/>
    public bool ReplaceTextFromSpellingChecker
    {
        get => _autoCorrect?.ReplaceTextFromSpellingChecker ?? false;
        set { if (_autoCorrect != null) _autoCorrect.ReplaceTextFromSpellingChecker = value; }
    }

    /// <inheritdoc/>
    public bool TwoInitialCapsAutoAdd
    {
        get => _autoCorrect?.TwoInitialCapsAutoAdd ?? false;
        set { if (_autoCorrect != null) _autoCorrect.TwoInitialCapsAutoAdd = value; }
    }

    /// <inheritdoc/>
    public MsWord.TwoInitialCapsExceptions TwoInitialCapsExceptions => _autoCorrect?.TwoInitialCapsExceptions;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordAutoCorrect"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _autoCorrect != null)
        {
            Marshal.ReleaseComObject(_autoCorrect);
            _autoCorrect = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordAutoCorrect"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
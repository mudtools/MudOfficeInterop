//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 查找实现类
/// </summary>
internal class WordFind : IWordFind
{
    private MsWord.Find? _find;
    private bool _disposedValue;

    /// <inheritdoc/>
    public IWordApplication? Application => _find != null ? new WordApplication(_find.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _find?.Parent;

    public string FindText
    {
        get => _find != null ? _find.Text : "";
        set
        {
            if (_find != null)
                _find.Text = value;

        }
    }

    public int NoProofing
    {
        get => _find != null ? _find.NoProofing : 0;
        set
        {
            if (_find != null)
                _find.NoProofing = value;
        }
    }

    public string ReplaceWith
    {
        get => _find != null ? _find.Replacement.Text : "";
        set
        {
            if (_find != null)
                _find.Replacement.Text = value;
        }
    }

    public bool MatchControl
    {
        get => _find != null ? _find.MatchControl : false;
        set
        {
            if (_find != null)
                _find.MatchControl = value;
        }
    }

    public bool MatchCase
    {
        get => _find != null ? _find.MatchCase : false;
        set
        {
            if (_find != null)
                _find.MatchCase = value;
        }
    }

    public bool MatchWholeWord
    {
        get => _find != null ? _find.MatchWholeWord : false;
        set
        {
            if (_find != null)
                _find.MatchWholeWord = value;
        }
    }

    public bool MatchWildcards
    {
        get => _find != null ? _find.MatchWildcards : false;
        set
        {
            if (_find != null)
                _find.MatchWildcards = value;
        }
    }

    public bool Format
    {
        get => _find != null ? _find.Format : false;
        set
        {
            if (_find != null)
                _find.Format = value;
        }
    }

    public bool MatchPrefix
    {
        get => _find != null ? _find.MatchPrefix : false;
        set
        {
            if (_find != null)
                _find.MatchPrefix = value;
        }
    }

    public bool MatchSuffix
    {
        get => _find != null ? _find.MatchSuffix : false;
        set
        {
            if (_find != null)
                _find.MatchSuffix = value;
        }
    }
    public bool MatchPhrase
    {
        get => _find != null ? _find.MatchPhrase : false;
        set
        {
            if (_find != null)
                _find.MatchPhrase = value;
        }
    }

    public bool MatchSoundsLike
    {
        get => _find != null ? _find.MatchSoundsLike : false;
        set
        {
            if (_find != null)
                _find.MatchSoundsLike = value;
        }
    }

    public bool MatchAllWordForms
    {
        get => _find != null ? _find.MatchAllWordForms : false;
        set
        {
            if (_find != null)
                _find.MatchAllWordForms = value;
        }
    }

    public bool MatchByte
    {
        get => _find != null ? _find.MatchByte : false;
        set
        {
            if (_find != null)
                _find.MatchByte = value;
        }
    }

    public bool MatchFuzzy
    {
        get => _find != null ? _find.MatchFuzzy : false;
        set
        {
            if (_find != null)
                _find.MatchFuzzy = value;
        }
    }

    public bool Found
    {
        get => _find != null ? _find.Found : false;
    }

    public bool IgnoreSpace
    {
        get => _find != null ? _find.IgnoreSpace : false;
        set
        {
            if (_find != null)
                _find.IgnoreSpace = value;
        }
    }


    public bool IgnorePunct
    {
        get => _find != null ? _find.IgnorePunct : false;
        set
        {
            if (_find != null)
                _find.IgnorePunct = value;
        }
    }

    public bool Forward
    {
        get => _find != null ? _find.Forward : false;
        set
        {
            if (_find != null)
                _find.Forward = value;
        }
    }



    public bool Highlight
    {
        get => _find != null ? _find.Highlight.ConvertToBool() : false;
        set
        {
            if (_find != null)
                _find.Highlight = value ? 1 : 0;
        }
    }

    public IWordFrame? Frame
    {
        get
        {
            if (_find != null)
                return new WordFrame(_find.Frame);
            return null;
        }
    }

    public IWordFont Font
    {
        get
        {
            if (_find != null)
                return new WordFont(_find.Font);
            return null;
        }
    }

    public IWordParagraphFormat ParagraphFormat
    {
        get
        {
            if (_find != null)
                return new WordParagraphFormat(_find.ParagraphFormat);
            return null;
        }

    }

    public IWordReplacement Replacement
    {
        get
        {
            if (_find != null)
                return new WordReplacement(_find.Replacement);
            return null;
        }
    }


    public WdFindWrap Wrap
    {
        get => _find != null ? _find.Wrap.EnumConvert(WdFindWrap.wdFindAsk) : WdFindWrap.wdFindAsk;
        set
        {
            if (_find != null)
                _find.Wrap = value.EnumConvert(MsWord.WdFindWrap.wdFindAsk);
        }
    }

    public WdLanguageID LanguageID
    {
        get => _find != null ? _find.LanguageID.EnumConvert(WdLanguageID.wdLanguageNone) : WdLanguageID.wdLanguageNone;
        set
        {
            if (_find != null)
                _find.LanguageID = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
        }
    }

    public WdBuiltinStyle Style
    {
        get => _find != null ? _find.get_Style().ObjectConvertEnum(WdBuiltinStyle.wdStyleNormal) : WdBuiltinStyle.wdStyleNormal;
        set
        {
            if (_find != null)
                _find.set_Style(value.EnumConvert(MsWord.WdBuiltinStyle.wdStyleNormal));
        }
    }


    internal WordFind(MsWord.Find find)
    {
        _find = find ?? throw new ArgumentNullException(nameof(find));
        _disposedValue = false;
    }

    public bool Execute(string? findText = null, bool? matchCase = false,
       bool? matchWholeWord = false, bool? matchWildcards = false,
       bool? matchSoundsLike = null, bool? matchAllWordForms = null,
       bool? forward = null, WdFindWrap? wrap = null, bool? format = null,
       string? replaceWith = null, WdReplace? replace = null)
    {
        if (_find == null)
            return false;

        try
        {
            return _find.Execute(
                    findText.ComArgsVal(),
                    matchCase.ComArgsVal(),
                    matchWholeWord.ComArgsVal(),
                    matchWildcards.ComArgsVal(),
                    matchSoundsLike.ComArgsVal(),
                    matchAllWordForms.ComArgsVal(),
                    forward.ComArgsVal(),
                    wrap.ComArgsConvert(d => d.EnumConvert(MsWord.WdFindWrap.wdFindAsk)).ComArgsVal(),
                    format.ComArgsVal(),
                    replaceWith.ComArgsVal(),
                    replace.ComArgsConvert(d => d.EnumConvert(MsWord.WdReplace.wdReplaceAll)).ComArgsVal(),
                    missing, missing, missing
            );
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to execute find operation.", ex);
        }
    }

    public bool HitHighlight(string? findText = null,
             WdColor? highlightColor = WdColor.wdColorYellow, WdColor? textColor = WdColor.wdColorBlack,
             bool? matchCase = false, bool? matchWholeWord = null,
             bool? matchPrefix = null, bool? matchSuffix = null,
             bool? matchPhrase = null, bool? matchWildcards = null,
             bool? matchSoundsLike = null, bool? matchAllWordForms = null,
             bool? matchByte = null, bool? matchFuzzy = null,
             bool? ignoreSpace = null, bool? ignorePunct = null)
    {
        if (_find == null)
            return false;
        try
        {
            return _find.HitHighlight(
                findText,
                highlightColor.ComArgsConvert(d => d.EnumConvert(MsWord.WdColor.wdColorYellow)).ComArgsVal(),
                textColor.ComArgsConvert(d => d.EnumConvert(MsWord.WdColor.wdColorBlack)).ComArgsVal(),
                matchCase.ComArgsVal(),
                matchWholeWord.ComArgsVal(),
                matchPrefix.ComArgsVal(),
                matchSuffix.ComArgsVal(),
                matchPhrase.ComArgsVal(),
                matchWildcards.ComArgsVal(),
                matchSoundsLike.ComArgsVal(),
                matchAllWordForms.ComArgsVal(),
                matchByte.ComArgsVal(),
                matchFuzzy.ComArgsVal(),
                missing,
                missing,
                missing,
                missing,
                ignoreSpace.ComArgsVal(),
                ignorePunct.ComArgsVal(),
                missing);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to execute hit highlight operation.", ex);
        }
    }

    public bool ExecuteReplace(WdReplace replace = WdReplace.wdReplaceAll)
    {
        if (_find == null)
            return false;
        return Execute(replace: replace);
    }

    public void ClearFormatting()
    {
        if (_find == null)
            return;
        try
        {
            _find.ClearFormatting();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear find formatting.", ex);
        }
    }

    public void ClearReplaceFormatting()
    {
        if (_find == null)
            return;
        try
        {
            _find.Replacement.ClearFormatting();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear replace formatting.", ex);
        }
    }

    public void ClearHitHighlight()
    {
        if (_find == null)
            return;
        try
        {
            _find.ClearHitHighlight();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear hit highlight.", ex);
        }
    }

    private static readonly object missing = System.Reflection.Missing.Value;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing)
        {
            // 释放域对象本身
            if (_find != null)
            {
                Marshal.ReleaseComObject(_find);
                _find = null;
            }
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
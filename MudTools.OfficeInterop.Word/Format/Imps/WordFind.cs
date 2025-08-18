//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    private readonly MsWord.Find _find;
    private bool _disposedValue;

    public string FindText
    {
        get => _find.Text;
        set => _find.Text = value;
    }

    public string ReplaceWith
    {
        get => _find.Replacement.Text;
        set => _find.Replacement.Text = value;
    }

    public bool MatchCase
    {
        get => _find.MatchCase;
        set => _find.MatchCase = value;
    }

    public bool MatchWholeWord
    {
        get => _find.MatchWholeWord;
        set => _find.MatchWholeWord = value;
    }

    public bool MatchWildcards
    {
        get => _find.MatchWildcards;
        set => _find.MatchWildcards = value;
    }

    public WdFindWrap Wrap
    {
        get => (WdFindWrap)_find.Wrap;
        set => _find.Wrap = (MsWord.WdFindWrap)value;
    }

    internal WordFind(MsWord.Find find)
    {
        _find = find ?? throw new ArgumentNullException(nameof(find));
        _disposedValue = false;
    }

    public bool Execute()
    {
        try
        {
            return _find.Execute();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to execute find operation.", ex);
        }
    }

    public bool ExecuteReplace(int replace = 2)
    {
        try
        {
            var replaceOptionObj = (object)(MsWord.WdReplace)replace;
            object findText = missing;
            object matchCase = missing;
            object matchWholeWord = missing;
            object matchWildcards = missing;
            object matchSoundsLike = missing;
            object matchAllWordForms = missing;
            object forward = missing;
            object wrap = missing;
            object format = missing;
            object replaceWith = missing;
            object replaceOption = replaceOptionObj;
            object matchKashida = missing;
            object matchDiacritics = missing;
            object matchAlefHamza = missing;
            object matchControl = missing;

            return _find.Execute(
                ref findText, ref matchCase, ref matchWholeWord, ref matchWildcards,
                ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap,
                ref format, ref replaceWith, ref replaceOption, ref matchKashida,
                ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to execute find and replace operation.", ex);
        }
    }

    public void ClearFormatting()
    {
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
        try
        {
            _find.Replacement.ClearFormatting();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear replace formatting.", ex);
        }
    }

    private static readonly object missing = System.Reflection.Missing.Value;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
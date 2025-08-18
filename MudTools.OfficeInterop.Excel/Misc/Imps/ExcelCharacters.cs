//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelCharacters : IExcelCharacters
{
    private MsExcel.Characters _characters;
    private MsExcel.Range _rang;
    private bool _disposedValue;

    public object Parent => _characters.Parent;

    public IExcelApplication Application
    {
        get
        {
            var application = _characters?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public int Count => _characters.Count;

    public string Text
    {
        get => _characters.Text;
        set => _characters.Text = value;
    }


    public IExcelFont Font => new ExcelFont(_characters.Font);

    internal ExcelCharacters(MsExcel.Characters characters, MsExcel.Range rang)
    {
        _characters = characters ?? throw new ArgumentNullException(nameof(characters));
        _rang = rang;
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _characters.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除字符。", ex);
        }
    }

    public IExcelCharacters this[int? start, int? length]
    {
        get
        {
            // 检查对象是否已释放
            if (_disposedValue)
                throw new ObjectDisposedException(nameof(ExcelCharacters));

            // 验证参数范围
            if (start < 1 || start > Count)
                throw new ArgumentOutOfRangeException("开始索引超出字符范围。");

            if (length < 0 || (start + length - 1) > Count)
                throw new ArgumentOutOfRangeException("截取长度超出字符范围。");

            try
            {
                // 创建子字符范围
                var subChars = _rang.Characters[start.ComArgsVal(), length.ComArgsVal()];
                return new ExcelCharacters(subChars, this._rang);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("字符内容截取失败:" + ex.Message, ex);
            }
        }
    }

    public IExcelCharacters Insert(string text)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("文本不能为空。", nameof(text));

        try
        {
            var result = _characters.Insert(text) as MsExcel.Characters;
            return result != null ? new ExcelCharacters(result, this._rang) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法插入文本: {text}", ex);
        }
    }


    public int Find(string what, int after = 1, bool matchCase = false, bool matchWholeWord = false)
    {
        if (string.IsNullOrEmpty(what))
            throw new ArgumentException("查找文本不能为空。", nameof(what));

        if (after < 1 || after > Count)
            throw new ArgumentOutOfRangeException(nameof(after));

        try
        {
            var text = Text;
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            if (matchWholeWord)
            {
                return FindWholeWord(text, what, after, comparison);
            }
            else
            {
                var index = text.IndexOf(what, after - 1, comparison);
                return index >= 0 ? index + 1 : 0;
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法查找文本: {what}", ex);
        }
    }

    public int Replace(string what, string replacement, bool matchCase = false, bool matchWholeWord = false)
    {
        if (string.IsNullOrEmpty(what))
            throw new ArgumentException("要替换的文本不能为空。", nameof(what));

        try
        {
            var originalText = Text;
            var newText = string.Empty;
            var count = 0;

            if (matchWholeWord)
            {
                newText = ReplaceWholeWord(originalText, what, replacement, matchCase, ref count);
            }
            else
            {
                var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                newText = originalText.Replace(what, replacement, comparison, ref count);
            }

            if (count > 0)
            {
                Text = newText;
            }

            return count;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法替换文本。原文本: {what}, 替换文本: {replacement}", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _characters != null)
        {
            try
            {
                Marshal.ReleaseComObject(_characters);
            }
            catch { }
            _characters = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private int FindWholeWord(string text, string word, int start, StringComparison comparison)
    {
        var index = text.IndexOf(word, start - 1, comparison);
        while (index >= 0)
        {
            var beforeValid = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
            var afterValid = index + word.Length >= text.Length || !char.IsLetterOrDigit(text[index + word.Length]);

            if (beforeValid && afterValid)
            {
                return index + 1;
            }

            index = text.IndexOf(word, index + 1, comparison);
        }
        return 0;
    }

    private string ReplaceWholeWord(string text, string word, string replacement, bool matchCase, ref int count)
    {
        var result = text;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = FindWholeWord(result, word, 1, comparison);

        while (index > 0)
        {
            var beforeValid = index == 1 || !char.IsLetterOrDigit(result[index - 2]);
            var afterValid = index + word.Length - 1 >= result.Length || !char.IsLetterOrDigit(result[index + word.Length - 1]);

            if (beforeValid && afterValid)
            {
                result = result.Substring(0, index - 1) + replacement + result.Substring(index + word.Length - 1);
                count++;
            }

            index = FindWholeWord(result, word, index + replacement.Length, comparison);
        }

        return result;
    }
}

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Words 的封装实现类。
/// </summary>
internal class WordWords : IWordWords
{
    private MsWord.Words _words;
    private bool _disposedValue;

    internal WordWords(MsWord.Words words)
    {
        _words = words ?? throw new ArgumentNullException(nameof(words));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _words != null ? new WordApplication(_words.Application) : null;

    /// <inheritdoc/>
    public object Parent => _words?.Parent;

    /// <inheritdoc/>
    public int Count => _words?.Count ?? 0;

    /// <inheritdoc/>
    public IWordRange First => _words?.First != null ? new WordRange(_words.First) : null;

    /// <inheritdoc/>
    public IWordRange Last => _words?.Last != null ? new WordRange(_words.Last) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordRange this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comRange = _words[index];
                return new WordRange(comRange);
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordRange GetRange(int start, int end)
    {
        if (_words == null) return null;

        // 确保索引在有效范围内
        int validStart = Math.Max(1, Math.Min(start, Count));
        int validEnd = Math.Max(validStart, Math.Min(end, Count));

        try
        {
            // 在 Word 中，Words 集合的范围操作需要通过父 Range 来实现
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                var startWord = _words[validStart];
                var endWord = _words[validEnd];

                var newRange = parentRange.Duplicate;
                newRange.SetRange(startWord.Start, endWord.End);
                return new WordRange(newRange);
            }
        }
        catch
        {
            // 如果无法创建范围，返回 null
        }

        return null;
    }

    /// <inheritdoc/>
    public IWordRange Find(string findText, bool forward = true, WdFindWrap wrap = WdFindWrap.wdFindStop,
                          bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false,
                          bool matchSoundsLike = false, bool matchAllWordForms = false)
    {
        if (_words == null || string.IsNullOrEmpty(findText)) return null;

        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                var findObj = parentRange.Find;
                findObj.ClearFormatting();
                findObj.Text = findText;
                findObj.Forward = forward;
                findObj.Wrap = (MsWord.WdFindWrap)(int)wrap;
                findObj.MatchCase = matchCase;
                findObj.MatchWholeWord = matchWholeWord;
                findObj.MatchWildcards = matchWildcards;
                findObj.MatchSoundsLike = matchSoundsLike;
                findObj.MatchAllWordForms = matchAllWordForms;

                if (findObj.Execute())
                {
                    return new WordRange(parentRange);
                }
            }
        }
        catch
        {
            // 查找失败返回 null
        }

        return null;
    }

    /// <inheritdoc/>
    public int Replace(string findText, string replaceText, bool replaceAll = false,
                      bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false,
                      bool matchSoundsLike = false, bool matchAllWordForms = false)
    {
        if (_words == null || string.IsNullOrEmpty(findText)) return 0;

        int replaceCount = 0;
        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                var findObj = parentRange.Find;
                findObj.ClearFormatting();
                findObj.Replacement.ClearFormatting();
                findObj.Text = findText;
                findObj.Replacement.Text = replaceText ?? string.Empty;
                findObj.Forward = true;
                findObj.Wrap = MsWord.WdFindWrap.wdFindStop;
                findObj.Format = false;
                findObj.MatchCase = matchCase;
                findObj.MatchWholeWord = matchWholeWord;
                findObj.MatchWildcards = matchWildcards;
                findObj.MatchSoundsLike = matchSoundsLike;
                findObj.MatchAllWordForms = matchAllWordForms;

                if (replaceAll)
                {
                    // 替换所有匹配项
                    while (findObj.Execute(Replace: MsWord.WdReplace.wdReplaceAll))
                    {
                        replaceCount++;
                    }
                }
                else
                {
                    // 只替换第一个匹配项
                    if (findObj.Execute(Replace: MsWord.WdReplace.wdReplaceOne))
                    {
                        replaceCount = 1;
                    }
                }
            }
        }
        catch
        {
            // 替换失败返回 0
        }

        return replaceCount;
    }

    /// <inheritdoc/>
    public IWordRange Add(string text)
    {
        if (_words == null || string.IsNullOrEmpty(text)) return null;

        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                // 移动到末尾并插入文本
                var endRange = parentRange.Duplicate;
                endRange.Collapse(MsWord.WdCollapseDirection.wdCollapseEnd);
                endRange.Text = text;

                // 返回新插入的文本范围
                var newRange = endRange.Duplicate;
                newRange.SetRange(endRange.End - text.Length, endRange.End);
                return new WordRange(newRange);
            }
        }
        catch
        {
            // 添加失败返回 null
        }

        return null;
    }

    /// <inheritdoc/>
    public IWordRange Insert(int index, string text)
    {
        if (_words == null || string.IsNullOrEmpty(text) || index < 1 || index > Count + 1) return null;

        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                // 计算插入位置
                int insertPosition = index <= Count ? _words[index].Start : parentRange.End;
                var insertRange = parentRange.Duplicate;
                insertRange.SetRange(insertPosition, insertPosition);
                insertRange.Text = text;

                // 返回新插入的文本范围
                var newRange = insertRange.Duplicate;
                newRange.SetRange(insertPosition, insertPosition + text.Length);
                return new WordRange(newRange);
            }
        }
        catch
        {
            // 插入失败返回 null
        }

        return null;
    }

    /// <inheritdoc/>
    public void Delete(int start, int count)
    {
        if (_words == null || start < 1 || count <= 0) return;

        try
        {
            int deleteStart = Math.Max(1, start);
            int deleteEnd = Math.Min(Count, deleteStart + count - 1);

            if (deleteStart <= Count && deleteEnd >= deleteStart)
            {
                var startWord = _words[deleteStart];
                var endWord = _words[deleteEnd];

                var deleteRange = startWord.Duplicate;
                deleteRange.SetRange(startWord.Start, endWord.End);
                deleteRange.Delete();
            }
        }
        catch
        {
            // 删除失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_words == null) return;

        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            parentRange?.Delete();
        }
        catch
        {
            // 清除失败忽略异常
        }
    }

    /// <inheritdoc/>
    public List<int> GetIndexes()
    {
        var indexes = new List<int>();
        for (int i = 1; i <= Count; i++)
        {
            indexes.Add(i);
        }
        return indexes;
    }

    /// <inheritdoc/>
    public string GetText()
    {
        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            return parentRange?.Text ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_words == null) return;

        try
        {
            var parentRange = _words.Parent as MsWord.Range;
            if (parentRange != null)
            {
                parentRange.Text = text ?? string.Empty;
            }
        }
        catch
        {
            // 设置文本失败忽略异常
        }
    }

    /// <inheritdoc/>
    public IWordRange Substring(int startIndex, int length)
    {
        if (startIndex < 1 || length <= 0 || startIndex + length - 1 > Count) return null;

        return GetRange(startIndex, startIndex + length - 1);
    }

    /// <inheritdoc/>
    public bool Contains(string text, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(text)) return false;

        try
        {
            var foundRange = Find(text, matchCase: matchCase);
            return foundRange != null;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public int GetCharacterCount(int index)
    {
        if (index < 1 || index > Count) return 0;

        try
        {
            var word = _words[index];
            return word?.Characters?.Count ?? 0;
        }
        catch
        {
            return 0;
        }
    }

    /// <inheritdoc/>
    public int GetSentenceCount(int startIndex, int length)
    {
        if (startIndex < 1 || length <= 0 || startIndex + length - 1 > Count) return 0;

        try
        {
            var startWord = _words[startIndex];
            var endWord = _words[Math.Min(startIndex + length - 1, Count)];

            var range = startWord.Duplicate;
            range.SetRange(startWord.Start, endWord.End);
            return range?.Sentences?.Count ?? 0;
        }
        catch
        {
            return 0;
        }
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordRange> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _words != null)
        {
            Marshal.ReleaseComObject(_words);
            _words = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
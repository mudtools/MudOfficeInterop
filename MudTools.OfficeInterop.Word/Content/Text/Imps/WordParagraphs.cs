//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Paragraphs 的封装实现类。
/// </summary>
internal class WordParagraphs : IWordParagraphs
{
    private MsWord.Paragraphs _paragraphs;
    private bool _disposedValue;

    internal WordParagraphs(MsWord.Paragraphs paragraphs)
    {
        _paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _paragraphs != null ? new WordApplication(_paragraphs.Application) : null;

    /// <inheritdoc/>
    public object Parent => _paragraphs?.Parent;

    /// <inheritdoc/>
    public int Count => _paragraphs?.Count ?? 0;

    /// <inheritdoc/>
    public IWordParagraph First => _paragraphs?.Count > 0 ? new WordParagraph(_paragraphs[1]) : null;

    /// <inheritdoc/>
    public IWordParagraph Last => _paragraphs?.Count > 0 ? new WordParagraph(_paragraphs[_paragraphs.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordParagraph this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comParagraph = _paragraphs[index];
                return new WordParagraph(comParagraph);
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
    public IWordParagraph Add(string text = null, int beforeParagraph = -1)
    {
        if (_paragraphs == null) return null;

        try
        {
            MsWord.Paragraph newParagraph;
            if (beforeParagraph > 0 && beforeParagraph <= Count)
            {
                // 在指定段落前添加
                var targetParagraph = _paragraphs[beforeParagraph];
                newParagraph = _paragraphs.Add(targetParagraph.Range);
            }
            else
            {
                // 在末尾添加
                newParagraph = _paragraphs.Add();
            }

            // 设置文本内容
            if (!string.IsNullOrEmpty(text))
            {
                newParagraph.Range.Text = text + "\r";
            }

            return new WordParagraph(newParagraph);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加段落。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordParagraph Insert(int index, string text = null)
    {
        if (_paragraphs == null || index < 1 || index > Count + 1) return null;

        try
        {
            MsWord.Paragraph newParagraph;
            if (index <= Count)
            {
                // 在指定位置插入
                var targetParagraph = _paragraphs[index];
                newParagraph = _paragraphs.Add(targetParagraph.Range);
            }
            else
            {
                // 在末尾添加
                newParagraph = _paragraphs.Add();
            }

            // 设置文本内容
            if (!string.IsNullOrEmpty(text))
            {
                newParagraph.Range.Text = text + "\r";
            }

            return new WordParagraph(newParagraph);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法插入段落。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        try
        {
            _paragraphs[index].Range.Delete();
        }
        catch
        {
            // 删除失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void DeleteRange(int startIndex, int count)
    {
        if (startIndex < 1 || count <= 0) return;

        try
        {
            int endIndex = Math.Min(startIndex + count - 1, Count);
            if (startIndex <= endIndex)
            {
                // 从后往前删除，避免索引变化
                for (int i = endIndex; i >= startIndex; i--)
                {
                    _paragraphs[i].Range.Delete();
                }
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
        if (_paragraphs == null) return;

        try
        {
            // 从后往前删除，避免索引变化
            for (int i = Count; i >= 1; i--)
            {
                _paragraphs[i].Range.Delete();
            }
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
    public int GetTotalCharacters()
    {
        int total = 0;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                total += _paragraphs[i].Range.Characters.Count;
            }
        }
        catch
        {
            // 统计失败返回 0
        }
        return total;
    }

    /// <inheritdoc/>
    public int GetTotalWords()
    {
        int total = 0;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                total += _paragraphs[i].Range.Words.Count;
            }
        }
        catch
        {
            // 统计失败返回 0
        }
        return total;
    }

    /// <inheritdoc/>
    public int GetTotalSentences()
    {
        int total = 0;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                total += _paragraphs[i].Range.Sentences.Count;
            }
        }
        catch
        {
            // 统计失败返回 0
        }
        return total;
    }

    /// <inheritdoc/>
    public void Sort(bool ascending = true)
    {
        if (_paragraphs == null) return;

        try
        {
            // 段落排序通常通过文档范围排序实现
            var parentRange = _paragraphs.Parent as MsWord.Range;
            if (parentRange != null)
            {
                parentRange.Sort(SortOrder: ascending ? MsWord.WdSortOrder.wdSortOrderAscending : MsWord.WdSortOrder.wdSortOrderDescending);
            }
        }
        catch
        {
            // 排序失败忽略异常
        }
    }

    /// <inheritdoc/>
    public List<IWordParagraph> FindByText(string text, bool matchCase = false, bool matchWholeWord = false)
    {
        var foundParagraphs = new List<IWordParagraph>();
        if (_paragraphs == null || string.IsNullOrEmpty(text)) return foundParagraphs;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var paragraph = _paragraphs[i];
                if (paragraph?.Range?.Text != null)
                {
                    string paragraphText = paragraph.Range.Text.TrimEnd('\r');
                    bool isMatch = false;

                    if (matchWholeWord)
                    {
                        // 匹配整个单词
                        string[] words = paragraphText.Split([' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                        isMatch = matchCase ?
                            Array.Exists(words, w => w.Equals(text)) :
                            Array.Exists(words, w => w.Equals(text, StringComparison.OrdinalIgnoreCase));
                    }
                    else
                    {
                        // 部分匹配
                        isMatch = matchCase ?
                            paragraphText.Contains(text) :
                            paragraphText.ToLower().Contains(text.ToLower());
                    }

                    if (isMatch)
                    {
                        foundParagraphs.Add(new WordParagraph(paragraph));
                    }
                }
            }
        }
        catch
        {
            // 查找失败返回空列表
        }

        return foundParagraphs;
    }

    /// <inheritdoc/>
    public int ReplaceAllText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (_paragraphs == null || string.IsNullOrEmpty(findText)) return 0;

        int totalReplacements = 0;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var paragraph = new WordParagraph(_paragraphs[i]);
                totalReplacements += paragraph.ReplaceText(findText, replaceText, matchCase, matchWholeWord);
                paragraph.Dispose();
            }
        }
        catch
        {
            // 替换失败返回已执行的替换次数
        }

        return totalReplacements;
    }

    /// <inheritdoc/>
    public List<IWordParagraph> GetRange(int startIndex, int endIndex)
    {
        var paragraphs = new List<IWordParagraph>();
        if (_paragraphs == null) return paragraphs;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        try
        {
            for (int i = validStart; i <= validEnd; i++)
            {
                paragraphs.Add(new WordParagraph(_paragraphs[i]));
            }
        }
        catch
        {
            // 获取范围失败返回已获取的段落
        }

        return paragraphs;
    }

    /// <inheritdoc/>
    public List<IWordParagraph> GetHeadings()
    {
        var headings = new List<IWordParagraph>();
        if (_paragraphs == null) return headings;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var paragraph = _paragraphs[i];
                if (paragraph?.OutlineLevel >= MsWord.WdOutlineLevel.wdOutlineLevel1 &&
                    paragraph?.OutlineLevel <= MsWord.WdOutlineLevel.wdOutlineLevel9)
                {
                    headings.Add(new WordParagraph(paragraph));
                }
            }
        }
        catch
        {
            // 获取标题失败返回已获取的标题
        }

        return headings;
    }

    /// <inheritdoc/>
    public List<IWordParagraph> GetEmptyParagraphs()
    {
        var emptyParagraphs = new List<IWordParagraph>();
        if (_paragraphs == null) return emptyParagraphs;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var paragraph = _paragraphs[i];
                string text = paragraph?.Range?.Text?.TrimEnd('\r');
                if (string.IsNullOrEmpty(text))
                {
                    emptyParagraphs.Add(new WordParagraph(paragraph));
                }
            }
        }
        catch
        {
            // 获取空段落失败返回已获取的段落
        }

        return emptyParagraphs;
    }

    /// <inheritdoc/>
    public void SetFormatForAll(WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft,
                               float leftIndent = 0, float lineSpacing = 1.0f)
    {
        if (_paragraphs == null) return;

        for (int i = 1; i <= Count; i++)
        {
            var para = _paragraphs[i];
            para.Alignment = (MsWord.WdParagraphAlignment)(int)alignment;
            para.LeftIndent = leftIndent;
            para.LineSpacing = lineSpacing;
        }
    }

    /// <inheritdoc/>
    public (int MinLength, int MaxLength, double AverageLength) GetLengthStatistics()
    {
        if (_paragraphs == null || Count == 0)
            return (0, 0, 0.0);

        try
        {
            int minLength = int.MaxValue;
            int maxLength = 0;
            int totalLength = 0;

            for (int i = 1; i <= Count; i++)
            {
                int length = _paragraphs[i].Range.Characters.Count;
                minLength = Math.Min(minLength, length);
                maxLength = Math.Max(maxLength, length);
                totalLength += length;
            }

            double averageLength = (double)totalLength / Count;
            return (minLength, maxLength, averageLength);
        }
        catch
        {
            return (0, 0, 0.0);
        }
    }
    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordParagraph> GetEnumerator()
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

        if (disposing && _paragraphs != null)
        {
            Marshal.ReleaseComObject(_paragraphs);
            _paragraphs = null;
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
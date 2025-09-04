//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Footnotes 的封装实现类。
/// </summary>
internal class WordFootnotes : IWordFootnotes
{
    private MsWord.Footnotes _footnotes;
    private bool _disposedValue;

    internal WordFootnotes(MsWord.Footnotes footnotes)
    {
        _footnotes = footnotes ?? throw new ArgumentNullException(nameof(footnotes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _footnotes != null ? new WordApplication(_footnotes.Application) : null;

    /// <inheritdoc/>
    public object Parent => _footnotes?.Parent;

    /// <inheritdoc/>
    public int Count => _footnotes?.Count ?? 0;

    /// <inheritdoc/>
    public IWordFootnote? First => _footnotes?.Count > 0 ? new WordFootnote(_footnotes[1]) : null;

    /// <inheritdoc/>
    public IWordFootnote? Last => _footnotes?.Count > 0 ? new WordFootnote(_footnotes[_footnotes.Count]) : null;

    /// <inheritdoc/>
    public IWordRange? Separator => _footnotes?.Separator != null ? new WordRange(_footnotes.Separator) : null;

    /// <inheritdoc/>
    public IWordRange? ContinuationSeparator => _footnotes?.ContinuationSeparator != null ? new WordRange(_footnotes.ContinuationSeparator) : null;

    /// <inheritdoc/>
    public IWordRange? ContinuationNotice => _footnotes?.ContinuationNotice != null ? new WordRange(_footnotes.ContinuationNotice) : null;

    /// <inheritdoc/>
    public WdNoteNumberStyle NumberStyle
    {
        get => _footnotes?.NumberStyle != null ? (WdNoteNumberStyle)(int)_footnotes?.NumberStyle : WdNoteNumberStyle.wdNoteNumberStyleArabic;
        set
        {
            if (_footnotes != null) _footnotes.NumberStyle = (MsWord.WdNoteNumberStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public int StartingNumber
    {
        get => _footnotes?.StartingNumber ?? 1;
        set
        {
            if (_footnotes != null)
                _footnotes.StartingNumber = value;
        }
    }

    /// <inheritdoc/>
    public WdNumberingRule NumberingRule
    {
        get => _footnotes?.NumberingRule != null ? (WdNumberingRule)(int)_footnotes?.NumberingRule : WdNumberingRule.wdRestartContinuous;
        set
        {
            if (_footnotes != null) _footnotes.NumberingRule = (MsWord.WdNumberingRule)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdFootnoteLocation Location
    {
        get => _footnotes?.Location != null ? (WdFootnoteLocation)(int)_footnotes?.Location : WdFootnoteLocation.wdBeneathText;
        set
        {
            if (_footnotes != null) _footnotes.Location = (MsWord.WdFootnoteLocation)(int)value;
        }
    }

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordFootnote this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comFootnote = _footnotes[index];
                return new WordFootnote(comFootnote);
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
    public IWordFootnote Add(IWordRange range, string referenceText = null, string noteText = null)
    {
        if (_footnotes == null || range == null) return null;

        try
        {
            // 获取原始 Range 对象
            var comRange = (range as WordRange)?._range;
            if (comRange != null)
            {
                MsWord.Range referenceRange = null;
                if (!string.IsNullOrEmpty(referenceText))
                {
                    referenceRange = comRange.Duplicate;
                    referenceRange.Text = referenceText;
                }

                var newFootnote = _footnotes.Add(comRange, referenceRange);

                // 设置脚注文本
                if (!string.IsNullOrEmpty(noteText))
                {
                    newFootnote.Range.Text = noteText;
                }

                return new WordFootnote(newFootnote);
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加脚注。", ex);
        }

        return null;
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        try
        {
            _footnotes[index].Delete();
        }
        catch
        {
            // 删除失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_footnotes == null) return;

        try
        {
            // 从后往前删除，避免索引变化
            for (int i = Count; i >= 1; i--)
            {
                _footnotes[i].Delete();
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
    public void Renumber()
    {
        if (_footnotes == null) return;

        // 重新编号通常通过更新域来实现
        var parentDocument = _footnotes.Parent as MsWord.Document;
        if (parentDocument != null)
        {
            parentDocument.Fields.Update();
        }
    }

    /// <inheritdoc/>
    public int CountInRange(IWordRange range)
    {
        if (_footnotes == null || range == null) return 0;

        int count = 0;
        var comRange = (range as WordRange)?._range;
        if (comRange != null)
        {
            for (int i = 1; i <= Count; i++)
            {
                var footnote = _footnotes[i];
                if (footnote != null && footnote.Reference != null)
                {
                    // 检查脚注引用是否在指定范围内
                    if (footnote.Reference.Start >= comRange.Start && footnote.Reference.End <= comRange.End)
                    {
                        count++;
                    }
                }
            }
        }

        return count;
    }

    /// <inheritdoc/>
    public List<IWordFootnote> FindByText(string text, bool matchCase = false)
    {
        var foundFootnotes = new List<IWordFootnote>();
        if (_footnotes == null || string.IsNullOrEmpty(text)) return foundFootnotes;

        for (int i = 1; i <= Count; i++)
        {
            var footnote = _footnotes[i];
            if (footnote?.Range?.Text != null)
            {
                string footnoteText = footnote.Range.Text;
                bool isMatch = matchCase ?
                    footnoteText.Contains(text) :
                    footnoteText.ToLower().Contains(text.ToLower());

                if (isMatch)
                {
                    foundFootnotes.Add(new WordFootnote(footnote));
                }
            }
        }

        return foundFootnotes;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordFootnote> GetEnumerator()
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

        if (disposing && _footnotes != null)
        {
            Marshal.ReleaseComObject(_footnotes);
            _footnotes = null;
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
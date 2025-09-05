//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Endnotes 的封装实现类。
/// </summary>
internal class WordEndnotes : IWordEndnotes
{
    private MsWord.Endnotes _endnotes;
    private bool _disposedValue;

    internal WordEndnotes(MsWord.Endnotes endnotes)
    {
        _endnotes = endnotes ?? throw new ArgumentNullException(nameof(endnotes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _endnotes != null ? new WordApplication(_endnotes.Application) : null;

    /// <inheritdoc/>
    public object Parent => _endnotes?.Parent;

    /// <inheritdoc/>
    public int Count => _endnotes?.Count ?? 0;

    /// <inheritdoc/>
    public IWordEndnote First => _endnotes?.Count > 0 ? new WordEndnote(_endnotes[1]) : null;

    /// <inheritdoc/>
    public IWordEndnote Last => _endnotes?.Count > 0 ? new WordEndnote(_endnotes[_endnotes.Count]) : null;


    /// <inheritdoc/>
    public IWordRange Separator => _endnotes?.Separator != null ? new WordRange(_endnotes.Separator) : null;

    /// <inheritdoc/>
    public IWordRange ContinuationSeparator => _endnotes?.ContinuationSeparator != null ? new WordRange(_endnotes.ContinuationSeparator) : null;

    /// <inheritdoc/>
    public IWordRange ContinuationNotice => _endnotes?.ContinuationNotice != null ? new WordRange(_endnotes.ContinuationNotice) : null;


    /// <inheritdoc/>
    public WdNoteNumberStyle NumberStyle
    {
        get => _endnotes?.NumberStyle != null ? (WdNoteNumberStyle)(int)_endnotes?.NumberStyle : WdNoteNumberStyle.wdNoteNumberStyleArabic;
        set
        {
            if (_endnotes != null) _endnotes.NumberStyle = (MsWord.WdNoteNumberStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public int StartingNumber
    {
        get => _endnotes?.StartingNumber ?? 1;
        set
        {
            if (_endnotes != null)
                _endnotes.StartingNumber = value;
        }
    }

    /// <inheritdoc/>
    public WdNumberingRule NumberingRule
    {
        get => _endnotes?.NumberingRule != null ? (WdNumberingRule)(int)_endnotes?.NumberingRule : WdNumberingRule.wdRestartContinuous;
        set
        {
            if (_endnotes != null) _endnotes.NumberingRule = (MsWord.WdNumberingRule)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdEndnoteLocation Location
    {
        get => _endnotes?.Location != null ? (WdEndnoteLocation)(int)_endnotes?.Location : WdEndnoteLocation.wdEndOfSection;
        set
        {
            if (_endnotes != null) _endnotes.Location = (MsWord.WdEndnoteLocation)(int)value;
        }
    }

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordEndnote this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comEndnote = _endnotes[index];
                return new WordEndnote(comEndnote);
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
    public IWordEndnote Add(IWordRange range, string referenceText = null, string noteText = null)
    {
        if (_endnotes == null || range == null) return null;

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

                var newEndnote = _endnotes.Add(comRange, referenceRange);

                // 设置尾注文本
                if (!string.IsNullOrEmpty(noteText))
                {
                    newEndnote.Range.Text = noteText;
                }

                return new WordEndnote(newEndnote);
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加尾注。", ex);
        }

        return null;
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        try
        {
            _endnotes[index].Delete();
        }
        catch
        {
            // 删除失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_endnotes == null) return;

        try
        {
            // 从后往前删除，避免索引变化
            for (int i = Count; i >= 1; i--)
            {
                _endnotes[i].Delete();
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
        if (_endnotes == null) return;

        try
        {
            // 重新编号通常通过更新域来实现
            var parentDocument = _endnotes.Parent as MsWord.Document;
            if (parentDocument != null)
            {
                parentDocument.Fields.Update();
            }
        }
        catch
        {
            // 重新编号失败忽略异常
        }
    }

    /// <inheritdoc/>
    public int CountInRange(IWordRange range)
    {
        if (_endnotes == null || range == null) return 0;

        int count = 0;
        try
        {
            var comRange = (range as WordRange)?._range;
            if (comRange != null)
            {
                for (int i = 1; i <= Count; i++)
                {
                    var endnote = _endnotes[i];
                    if (endnote != null && endnote.Reference != null)
                    {
                        // 检查尾注引用是否在指定范围内
                        if (endnote.Reference.Start >= comRange.Start && endnote.Reference.End <= comRange.End)
                        {
                            count++;
                        }
                    }
                }
            }
        }
        catch
        {
            // 统计失败返回 0
        }

        return count;
    }

    /// <inheritdoc/>
    public List<IWordEndnote> FindByText(string text, bool matchCase = false)
    {
        var foundEndnotes = new List<IWordEndnote>();
        if (_endnotes == null || string.IsNullOrEmpty(text)) return foundEndnotes;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var endnote = _endnotes[i];
                if (endnote?.Range?.Text != null)
                {
                    string endnoteText = endnote.Range.Text;
                    bool isMatch = matchCase ?
                        endnoteText.Contains(text) :
                        endnoteText.ToLower().Contains(text.ToLower());

                    if (isMatch)
                    {
                        foundEndnotes.Add(new WordEndnote(endnote));
                    }
                }
            }
        }
        catch
        {
            // 查找失败返回空列表
        }

        return foundEndnotes;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordEndnote> GetEnumerator()
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

        if (disposing && _endnotes != null)
        {
            Marshal.ReleaseComObject(_endnotes);
            _endnotes = null;
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
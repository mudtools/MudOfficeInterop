//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Comments 的封装实现类。
/// </summary>
internal class WordComments : IWordComments
{
    private MsWord.Comments _comments;
    private bool _disposedValue;

    internal WordComments(MsWord.Comments comments)
    {
        _comments = comments ?? throw new ArgumentNullException(nameof(comments));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _comments != null ? new WordApplication(_comments.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _comments?.Parent;

    /// <inheritdoc/>
    public int Count => _comments?.Count ?? 0;

    /// <inheritdoc/>
    public IWordComment First => _comments?.Count > 0 ? new WordComment(_comments[1]) : null;

    /// <inheritdoc/>
    public IWordComment Last => _comments?.Count > 0 ? new WordComment(_comments[_comments.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordComment this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comComment = _comments[index];
                return new WordComment(comComment);
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
    public IWordComment Add(IWordRange range, string text = null, string author = null)
    {
        if (_comments == null || range == null) return null;

        try
        {
            // 获取原始 Range 对象
            var comRange = (range as WordRange)?.InternalComObject;
            if (comRange != null)
            {
                var newComment = _comments.Add(comRange, text ?? string.Empty);

                // 设置作者
                if (!string.IsNullOrEmpty(author))
                {
                    newComment.Author = author;
                }

                return new WordComment(newComment);
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加批注。", ex);
        }

        return null;
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        _comments[index].Delete();
    }

    /// <inheritdoc/>
    public void DeleteRange(int startIndex, int count)
    {
        if (startIndex < 1 || count <= 0) return;

        int endIndex = Math.Min(startIndex + count - 1, Count);
        if (startIndex <= endIndex)
        {
            // 从后往前删除，避免索引变化
            for (int i = endIndex; i >= startIndex; i--)
            {
                _comments[i].Delete();
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_comments == null) return;

        // 从后往前删除，避免索引变化
        for (int i = Count; i >= 1; i--)
        {
            _comments[i].Delete();
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
        for (int i = 1; i <= Count; i++)
        {
            total += _comments[i].Scope.Characters.Count;
        }
        return total;
    }

    /// <inheritdoc/>
    public int GetTotalWords()
    {
        int total = 0;
        for (int i = 1; i <= Count; i++)
        {
            total += _comments[i].Scope.Words.Count;
        }
        return total;
    }


    /// <inheritdoc/>
    public List<IWordComment> GetByDateRange(DateTime startDate, DateTime endDate)
    {
        var comments = new List<IWordComment>();
        if (_comments == null) return comments;

        for (int i = 1; i <= Count; i++)
        {
            var comment = _comments[i];
            if (comment?.Date >= startDate && comment?.Date <= endDate)
            {
                comments.Add(new WordComment(comment));
            }
        }

        return comments;
    }

    /// <inheritdoc/>
    public List<IWordComment> GetByAuthor(string author)
    {
        var comments = new List<IWordComment>();
        if (_comments == null || string.IsNullOrEmpty(author)) return comments;

        for (int i = 1; i <= Count; i++)
        {
            var comment = _comments[i];
            if (comment?.Author?.Equals(author, StringComparison.OrdinalIgnoreCase) == true)
            {
                comments.Add(new WordComment(comment));
            }
        }

        return comments;
    }

    /// <inheritdoc/>
    public List<IWordComment> FindByText(string text, bool matchCase = false, bool matchWholeWord = false)
    {
        var foundComments = new List<IWordComment>();
        if (_comments == null || string.IsNullOrEmpty(text)) return foundComments;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var comment = _comments[i];
                if (comment?.Scope?.Text != null)
                {
                    string commentText = comment.Scope.Text;
                    bool isMatch = false;

                    if (matchWholeWord)
                    {
                        // 匹配整个单词
                        string[] words = commentText.Split([' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                        isMatch = matchCase ?
                            Array.Exists(words, w => w.Equals(text)) :
                            Array.Exists(words, w => w.Equals(text, StringComparison.OrdinalIgnoreCase));
                    }
                    else
                    {
                        // 部分匹配
                        isMatch = matchCase ?
                            commentText.Contains(text) :
                            commentText.ToLower().Contains(text.ToLower());
                    }

                    if (isMatch)
                    {
                        foundComments.Add(new WordComment(comment));
                    }
                }
            }
        }
        catch
        {
            // 查找失败返回空列表
        }

        return foundComments;
    }

    /// <inheritdoc/>
    public int ReplaceAllText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (_comments == null || string.IsNullOrEmpty(findText)) return 0;

        int totalReplacements = 0;
        for (int i = 1; i <= Count; i++)
        {
            var comment = new WordComment(_comments[i]);
            totalReplacements += comment.ReplaceText(findText, replaceText, matchCase, matchWholeWord);
            comment.Dispose();
        }

        return totalReplacements;
    }

    /// <inheritdoc/>
    public List<IWordComment> GetRange(int startIndex, int endIndex)
    {
        var comments = new List<IWordComment>();
        if (_comments == null) return comments;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        for (int i = validStart; i <= validEnd; i++)
        {
            comments.Add(new WordComment(_comments[i]));
        }

        return comments;
    }
    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordComment> GetEnumerator()
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

        if (disposing && _comments != null)
        {
            Marshal.ReleaseComObject(_comments);
            _comments = null;
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
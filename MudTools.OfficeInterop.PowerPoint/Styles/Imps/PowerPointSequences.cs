//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 序列集合实现类
/// </summary>
internal class PowerPointSequences : IPowerPointSequences
{
    private readonly MsPowerPoint.TimeLine _timeLine;
    private bool _disposedValue;

    /// <summary>
    /// 获取序列数量
    /// </summary>
    public int Count
    {
        get
        {
            try
            {
                // 主序列 + 交互序列
                return 1 + (_timeLine?.InteractiveSequences?.Count ?? 0);
            }
            catch
            {
                return 0;
            }
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _timeLine?.Parent;

    /// <summary>
    /// 根据索引获取序列
    /// </summary>
    public IPowerPointSequence this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                if (index == 1)
                {
                    // 主序列
                    return _timeLine?.MainSequence != null ? new PowerPointSequence(_timeLine.MainSequence) : null;
                }
                else
                {
                    // 交互序列
                    var interactiveIndex = index - 1;
                    if (_timeLine?.InteractiveSequences != null && interactiveIndex <= _timeLine.InteractiveSequences.Count)
                    {
                        return new PowerPointSequence(_timeLine.InteractiveSequences[interactiveIndex]);
                    }
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get sequence at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="timeLine">COM TimeLine 对象</param>
    internal PowerPointSequences(MsPowerPoint.TimeLine timeLine)
    {
        _timeLine = timeLine;
        _disposedValue = false;
    }



    /// <summary>
    /// 添加序列
    /// </summary>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的序列</returns>
    public IPowerPointSequence Add(int index = -1)
    {
        try
        {
            if (_timeLine?.InteractiveSequences != null)
            {
                MsPowerPoint.Sequence sequence;
                if (index > 0 && index <= _timeLine.InteractiveSequences.Count + 1)
                {
                    sequence = _timeLine.InteractiveSequences.Add(index);
                }
                else
                {
                    sequence = _timeLine.InteractiveSequences.Add();
                }
                return new PowerPointSequence(sequence);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add sequence.", ex);
        }
    }


    /// <summary>
    /// 根据条件查找序列
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的序列列表</returns>
    public IEnumerable<IPowerPointSequence> Find(Func<IPowerPointSequence, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        var results = new List<IPowerPointSequence>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var sequence = this[i];
                    if (sequence != null && predicate(sequence))
                    {
                        results.Add(sequence);
                    }
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find sequences.", ex);
        }
        return results;
    }

    /// <summary>
    /// 获取主序列
    /// </summary>
    /// <returns>主序列</returns>
    public IPowerPointSequence GetMainSequence()
    {
        try
        {
            return _timeLine?.MainSequence != null ? new PowerPointSequence(_timeLine.MainSequence) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get main sequence.", ex);
        }
    }

    /// <summary>
    /// 获取交互序列
    /// </summary>
    /// <returns>交互序列</returns>
    public IEnumerable<IPowerPointSequence> GetInteractiveSequences()
    {
        var sequences = new List<IPowerPointSequence>();
        try
        {
            if (_timeLine?.InteractiveSequences != null)
            {
                foreach (MsPowerPoint.Sequence sequence in _timeLine.InteractiveSequences)
                {
                    sequences.Add(new PowerPointSequence(sequence));
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get interactive sequences.", ex);
        }
        return sequences;
    }

    /// <summary>
    /// 重新排序序列
    /// </summary>
    /// <param name="newOrder">新顺序数组</param>
    public void Reorder(int[] newOrder)
    {
        if (newOrder == null)
            throw new ArgumentNullException(nameof(newOrder));

        try
        {
            // 序列重新排序需要更复杂的实现
            throw new NotImplementedException("Reordering sequences is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reorder sequences.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    public IEnumerator<IPowerPointSequence> GetEnumerator()
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
}

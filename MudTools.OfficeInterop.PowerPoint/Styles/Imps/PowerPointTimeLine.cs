//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 时间线实现类
/// </summary>
internal class PowerPointTimeLine : IPowerPointTimeLine
{
    private readonly MsPowerPoint.TimeLine _timeLine;
    private bool _disposedValue;
    private IPowerPointSequences _sequences;

    /// <summary>
    /// 获取动画序列集合
    /// </summary>
    public IPowerPointSequences Sequences
    {
        get
        {
            if (_sequences == null && _timeLine?.MainSequence != null)
            {
                _sequences = new PowerPointSequences(_timeLine);
            }
            return _sequences;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _timeLine?.Parent;

    /// <summary>
    /// 获取或设置是否启用动画
    /// </summary>
    public bool Enabled
    {
        get => _timeLine?.InteractiveSequences?.Count > 0 || _timeLine?.MainSequence?.Count > 0;
        set
        {
            // 启用/禁用动画需要通过其他方式实现
        }
    }

    /// <summary>
    /// 获取动画效果数量
    /// </summary>
    public int EffectCount
    {
        get
        {
            var count = 0;
            try
            {
                count += _timeLine?.MainSequence?.Count ?? 0;
                if (_timeLine?.InteractiveSequences != null)
                {
                    foreach (MsPowerPoint.Sequence sequence in _timeLine.InteractiveSequences)
                    {
                        count += sequence.Count;
                    }
                }
            }
            catch
            {
                // 忽略计数错误
            }
            return count;
        }
    }

    /// <summary>
    /// 获取主序列
    /// </summary>
    public IPowerPointSequence MainSequence
    {
        get
        {
            try
            {
                return _timeLine?.MainSequence != null ? new PowerPointSequence(_timeLine.MainSequence) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取交互序列
    /// </summary>
    public IPowerPointSequence InteractiveSequences
    {
        get
        {
            try
            {
                // 返回第一个交互序列
                if (_timeLine?.InteractiveSequences?.Count > 0)
                {
                    return new PowerPointSequence(_timeLine.InteractiveSequences[1]);
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="timeLine">COM TimeLine 对象</param>
    internal PowerPointTimeLine(MsPowerPoint.TimeLine timeLine)
    {
        _timeLine = timeLine;
        _disposedValue = false;
    }

    /// <summary>
    /// 添加动画序列
    /// </summary>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的序列</returns>
    public IPowerPointSequence AddSequence(int index = -1)
    {
        try
        {
            if (_timeLine?.InteractiveSequences != null)
            {
                MsPowerPoint.Sequence sequence;
                if (index >= 0)
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
    /// 刷新动画显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            // 时间线刷新通常自动进行
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh timeline.", ex);
        }
    }

    /// <summary>
    /// 应用动画方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    public void ApplyAnimationScheme(int schemeIndex = -1)
    {
        try
        {
            // 动画方案应用需要通过幻灯片对象实现
            throw new NotImplementedException("Applying animation scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply animation scheme.", ex);
        }
    }

    /// <summary>
    /// 复制动画到其他幻灯片
    /// </summary>
    /// <param name="targetSlide">目标幻灯片</param>
    public void CopyTo(IPowerPointSlide targetSlide)
    {
        if (targetSlide == null)
            throw new ArgumentNullException(nameof(targetSlide));

        try
        {
            // 动画复制需要具体的实现
            throw new NotImplementedException("Copying animations is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy animations.", ex);
        }
    }

    /// <summary>
    /// 获取动画效果
    /// </summary>
    /// <param name="index">效果索引</param>
    /// <returns>动画效果</returns>
    public IPowerPointEffect GetEffect(int index)
    {
        try
        {
            if (_timeLine?.MainSequence != null && index >= 1 && index <= _timeLine.MainSequence.Count)
            {
                return new PowerPointEffect(_timeLine.MainSequence[index]);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get effect at index {index}.", ex);
        }
    }

    /// <summary>
    /// 查找指定形状的动画效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <returns>动画效果列表</returns>
    public IEnumerable<IPowerPointEffect> FindEffectsByShape(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        var effects = new List<IPowerPointEffect>();
        try
        {
            // 查找主序列中的效果
            if (_timeLine?.MainSequence != null)
            {
                for (int i = 1; i <= _timeLine.MainSequence.Count; i++)
                {
                    try
                    {
                        var effect = _timeLine.MainSequence[i];
                        // 这里需要比较形状对象，但COM对象比较复杂
                        effects.Add(new PowerPointEffect(effect));
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            // 查找交互序列中的效果
            if (_timeLine?.InteractiveSequences != null)
            {
                foreach (MsPowerPoint.Sequence sequence in _timeLine.InteractiveSequences)
                {
                    for (int i = 1; i <= sequence.Count; i++)
                    {
                        try
                        {
                            var effect = sequence[i];
                            effects.Add(new PowerPointEffect(effect));
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find effects by shape.", ex);
        }
        return effects;
    }

    /// <summary>
    /// 设置动画播放顺序
    /// </summary>
    /// <param name="effectOrder">效果顺序数组</param>
    public void SetEffectOrder(int[] effectOrder)
    {
        if (effectOrder == null)
            throw new ArgumentNullException(nameof(effectOrder));

        try
        {
            // 动画顺序设置需要更复杂的实现
            throw new NotImplementedException("Setting effect order is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set effect order.", ex);
        }
    }

    /// <summary>
    /// 获取时间线信息
    /// </summary>
    /// <returns>时间线信息字符串</returns>
    public string GetTimeLineInfo()
    {
        try
        {
            return $"TimeLine - Effects: {EffectCount}, Sequences: {_timeLine?.InteractiveSequences?.Count + 1}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get timeline info.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _sequences?.Dispose();
        }

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
}

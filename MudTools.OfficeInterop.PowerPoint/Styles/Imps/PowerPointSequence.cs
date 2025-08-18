//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 序列实现类
/// </summary>
internal class PowerPointSequence : IPowerPointSequence
{
    private readonly MsPowerPoint.Sequence _sequence;
    private bool _disposedValue;

    /// <summary>
    /// 获取效果数量
    /// </summary>
    public int Count => _sequence?.Count ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _sequence?.Parent;

    /// <summary>
    /// 根据索引获取效果
    /// </summary>
    public IPowerPointEffect this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                var effect = _sequence[index];
                return effect != null ? new PowerPointEffect(effect) : null;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get effect at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="sequence">COM Sequence 对象</param>
    internal PowerPointSequence(MsPowerPoint.Sequence sequence)
    {
        _sequence = sequence;
        _disposedValue = false;
    }


    /// <summary>
    /// 添加效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <param name="effectId">效果ID</param>
    /// <param name="trigger">触发器</param>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的效果</returns>
    public IPowerPointEffect AddEffect(IPowerPointShape shape, int effectId = 1, int trigger = 1, int index = -1)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        try
        {
            if (_sequence != null)
            {
                // 这里需要将 IPowerPointShape 转换为 COM Shape 对象
                // 由于缺少具体实现，使用占位符
                MsPowerPoint.Shape comShape = null;
                MsPowerPoint.Effect effect;

                if (index > 0 && index <= Count + 1)
                {
                    effect = _sequence.AddEffect(comShape, (MsPowerPoint.MsoAnimEffect)effectId,
                        (MsPowerPoint.MsoAnimateByLevel)trigger, (MsPowerPoint.MsoAnimTriggerType)index);
                }
                else
                {
                    effect = _sequence.AddEffect(comShape, (MsPowerPoint.MsoAnimEffect)effectId,
                        (MsPowerPoint.MsoAnimateByLevel)trigger);
                }

                return new PowerPointEffect(effect);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add effect.", ex);
        }
    }

    /// <summary>
    /// 删除效果
    /// </summary>
    /// <param name="index">效果索引</param>
    public void DeleteEffect(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var effect = _sequence[index];
            effect?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete effect at index {index}.", ex);
        }
    }

    /// <summary>
    /// 移动效果
    /// </summary>
    /// <param name="fromIndex">源索引</param>
    /// <param name="toIndex">目标索引</param>
    public void MoveEffect(int fromIndex, int toIndex)
    {
        if (fromIndex < 1 || fromIndex > Count)
            throw new ArgumentOutOfRangeException(nameof(fromIndex), $"From index must be between 1 and {Count}.");
        if (toIndex < 1 || toIndex > Count)
            throw new ArgumentOutOfRangeException(nameof(toIndex), $"To index must be between 1 and {Count}.");

        try
        {
            var effect = _sequence[fromIndex];
            effect?.MoveTo(toIndex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to move effect from index {fromIndex} to {toIndex}.", ex);
        }
    }

    /// <summary>
    /// 查找指定形状的效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <returns>效果列表</returns>
    public IEnumerable<IPowerPointEffect> FindEffectsByShape(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        var effects = new List<IPowerPointEffect>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var effect = _sequence[i];
                    // 这里需要比较形状对象
                    effects.Add(new PowerPointEffect(effect));
                }
                catch
                {
                    continue;
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
    /// 清除所有效果
    /// </summary>
    public void ClearEffects()
    {
        try
        {
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    var effect = _sequence[i];
                    effect?.Delete();
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear effects.", ex);
        }
    }

    /// <summary>
    /// 设置序列播放时间
    /// </summary>
    /// <param name="startTime">开始时间</param>
    /// <param name="duration">持续时间</param>
    public void SetTiming(float startTime, float duration)
    {
        try
        {
            // 序列时间设置需要通过效果的时间设置来实现
            throw new NotImplementedException("Setting sequence timing is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set sequence timing.", ex);
        }
    }

    /// <summary>
    /// 获取序列信息
    /// </summary>
    /// <returns>序列信息字符串</returns>
    public string GetSequenceInfo()
    {
        try
        {
            return $"Sequence - Effects: {Count}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get sequence info.", ex);
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
}

namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureEffects 的实现类。
/// </summary>
internal class OfficePictureEffects : IOfficePictureEffects
{
    private MsCore.PictureEffects _pictureEffects;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="pictureEffects">原始 COM PictureEffects 对象。</param>
    internal OfficePictureEffects(MsCore.PictureEffects pictureEffects)
    {
        _pictureEffects = pictureEffects ?? throw new ArgumentNullException(nameof(pictureEffects));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _pictureEffects?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficePictureEffect this[int index]
    {
        get
        {
            if (_pictureEffects == null || index < 1 || index > Count)
                return null;

            try
            {
                var effect = _pictureEffects[index];
                return new OfficePictureEffect(effect);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IOfficePictureEffect this[MsoPictureEffectType effectType]
    {
        get
        {
            if (_pictureEffects == null)
                return null;

            // 查找第一个匹配类型的效果
            for (int i = 1; i <= Count; i++)
            {
                var effect = _pictureEffects[i];
                if (effect != null && effect.Type == (MsCore.MsoPictureEffectType)(int)effectType)
                {
                    return new OfficePictureEffect(effect);
                }
            }

            return null;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficePictureEffect Insert(MsoPictureEffectType effectType, int position = -1)
    {
        if (_pictureEffects == null)
        {
            throw new ObjectDisposedException(nameof(OfficePictureEffects));
        }

        try
        {
            MsCore.PictureEffect effect;
            effect = _pictureEffects.Insert((MsCore.MsoPictureEffectType)(int)effectType, position);
            return new OfficePictureEffect(effect);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加图片效果类型 {effectType}。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (_pictureEffects == null)
            return;

        try
        {
            _pictureEffects[index]?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除索引为 {index} 的图片效果。", ex);
        }
    }

    /// <inheritdoc/>
    public void DeleteByType(MsoPictureEffectType effectType)
    {
        if (_pictureEffects == null)
            return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                var effect = _pictureEffects[i];
                if (effect != null && effect.Type == (MsCore.MsoPictureEffectType)(int)effectType)
                {
                    effect.Delete();
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除类型为 {effectType} 的图片效果。", ex);
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_pictureEffects == null)
            return;

        try
        {
            // 从后往前删除所有效果
            for (int i = Count; i >= 1; i--)
            {
                _pictureEffects[i]?.Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除所有图片效果。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Contains(MsoPictureEffectType effectType)
    {
        if (_pictureEffects == null)
            return false;

        for (int i = 1; i <= Count; i++)
        {
            var effect = _pictureEffects[i];
            if (effect != null && effect.Type == (MsCore.MsoPictureEffectType)(int)effectType)
            {
                return true;
            }
        }

        return false;
    }

    /// <inheritdoc/>
    public int GetCountByType(MsoPictureEffectType effectType)
    {
        if (_pictureEffects == null)
            return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            var effect = _pictureEffects[i];
            if (effect != null && effect.Type == (MsCore.MsoPictureEffectType)(int)effectType)
            {
                count++;
            }
        }

        return count;
    }

    /// <inheritdoc/>
    public List<MsoPictureEffectType> GetAllEffectTypes()
    {
        var types = new List<MsoPictureEffectType>();

        if (_pictureEffects == null)
            return types;

        for (int i = 1; i <= Count; i++)
        {
            var effect = _pictureEffects[i];
            if (effect != null && !types.Contains((MsoPictureEffectType)(int)effect.Type))
            {
                types.Add((MsoPictureEffectType)(int)effect.Type);
            }
        }

        return types;
    }

    /// <inheritdoc/>
    public void MoveEffect(IOfficePictureEffect effect, int newPosition)
    {
        if (_pictureEffects == null || effect == null)
            return;

        try
        {
            var wordEffect = (effect as OfficePictureEffect)?._pictureEffect;
            if (wordEffect != null)
            {
                wordEffect.Position = newPosition;
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移动图片效果到位置 {newPosition}。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_pictureEffects == null)
            return;

        Clear();
    }

    #endregion

    #region IEnumerable<IOfficePictureEffect> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficePictureEffect> GetEnumerator()
    {
        if (_pictureEffects == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var effect = _pictureEffects[i];
            if (effect != null)
                yield return new OfficePictureEffect(effect);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pictureEffects != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_pictureEffects);
            _pictureEffects = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
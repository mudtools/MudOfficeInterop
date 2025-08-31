namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureEffect 的实现类。
/// </summary>
internal class OfficePictureEffect : IOfficePictureEffect
{
    internal MsCore.PictureEffect _pictureEffect;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="pictureEffect">原始 COM PictureEffect 对象。</param>
    internal OfficePictureEffect(MsCore.PictureEffect pictureEffect)
    {
        _pictureEffect = pictureEffect ?? throw new ArgumentNullException(nameof(pictureEffect));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public MsoPictureEffectType Type
    {
        get => _pictureEffect?.Type != null ? (MsoPictureEffectType)(int)_pictureEffect.Type : MsoPictureEffectType.msoEffectNone;

    }

    /// <inheritdoc/>
    public int Position
    {
        get => _pictureEffect?.Position ?? 0;
        set
        {
            if (_pictureEffect != null)
                _pictureEffect.Position = value;
        }
    }

    /// <inheritdoc/>
    public bool Visible
    {
        get => _pictureEffect?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_pictureEffect != null)
                _pictureEffect.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }


    /// <inheritdoc/>
    public IOfficeEffectParameters EffectParameters => _pictureEffect?.EffectParameters != null ? new OfficeEffectParameters(_pictureEffect.EffectParameters) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _pictureEffect?.Delete();
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

        if (disposing)
        {
            // 释放效果参数集合
            if (_pictureEffect?.EffectParameters != null)
            {
                Marshal.ReleaseComObject(_pictureEffect.EffectParameters);
            }
            // 释放图片效果对象本身
            if (_pictureEffect != null)
            {
                Marshal.ReleaseComObject(_pictureEffect);
                _pictureEffect = null;
            }
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
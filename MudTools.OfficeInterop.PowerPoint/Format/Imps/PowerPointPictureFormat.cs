//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 图片格式实现类
/// </summary>
internal class PowerPointPictureFormat : IPowerPointPictureFormat
{
    private readonly MsPowerPoint.PictureFormat _pictureFormat;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置亮度
    /// </summary>
    public float Brightness
    {
        get => _pictureFormat?.Brightness ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Brightness = value;
        }
    }

    /// <summary>
    /// 获取或设置对比度
    /// </summary>
    public float Contrast
    {
        get => _pictureFormat?.Contrast ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Contrast = value;
        }
    }

    /// <summary>
    /// 获取或设置是否透明背景
    /// </summary>
    public bool TransparentBackground
    {
        get => _pictureFormat?.TransparentBackground == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.TransparentBackground = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置裁剪左边缘
    /// </summary>
    public float CropLeft
    {
        get => _pictureFormat?.CropLeft ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropLeft = value;
        }
    }

    /// <summary>
    /// 获取或设置裁剪右边缘
    /// </summary>
    public float CropRight
    {
        get => _pictureFormat?.CropRight ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropRight = value;
        }
    }

    /// <summary>
    /// 获取或设置裁剪上边缘
    /// </summary>
    public float CropTop
    {
        get => _pictureFormat?.CropTop ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropTop = value;
        }
    }

    /// <summary>
    /// 获取或设置裁剪下边缘
    /// </summary>
    public float CropBottom
    {
        get => _pictureFormat?.CropBottom ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropBottom = value;
        }
    }


    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _pictureFormat?.Parent;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="pictureFormat">COM PictureFormat 对象</param>
    internal PowerPointPictureFormat(MsPowerPoint.PictureFormat pictureFormat)
    {
        _pictureFormat = pictureFormat;
        _disposedValue = false;
    }

    /// <summary>
    /// 裁剪图片
    /// </summary>
    public void Crop()
    {
        try
        {
            // PictureFormat 本身没有直接的 Crop 方法
            // 裁剪通过设置 CropLeft, CropRight, CropTop, CropBottom 属性实现
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to crop picture.", ex);
        }
    }

    /// <summary>
    /// 重置图片格式
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.Brightness = 0;
                _pictureFormat.Contrast = 0;
                _pictureFormat.CropLeft = 0;
                _pictureFormat.CropRight = 0;
                _pictureFormat.CropTop = 0;
                _pictureFormat.CropBottom = 0;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset picture format.", ex);
        }
    }

    /// <summary>
    /// 设置裁剪区域
    /// </summary>
    /// <param name="left">左裁剪</param>
    /// <param name="right">右裁剪</param>
    /// <param name="top">上裁剪</param>
    /// <param name="bottom">下裁剪</param>
    public void SetCrop(float left, float right, float top, float bottom)
    {
        try
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.CropLeft = left;
                _pictureFormat.CropRight = right;
                _pictureFormat.CropTop = top;
                _pictureFormat.CropBottom = bottom;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set crop area.", ex);
        }
    }

    /// <summary>
    /// 应用图片样式
    /// </summary>
    /// <param name="styleIndex">样式索引</param>
    public void ApplyStyle(int styleIndex)
    {
        try
        {
            // 这需要通过父形状来应用样式
            throw new NotImplementedException("Applying picture style is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply picture style.", ex);
        }
    }

    /// <summary>
    /// 获取图片信息
    /// </summary>
    /// <returns>图片信息字符串</returns>
    public string GetPictureInfo()
    {
        try
        {
            return $"Brightness: {Brightness}, Contrast: {Contrast}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get picture info.", ex);
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

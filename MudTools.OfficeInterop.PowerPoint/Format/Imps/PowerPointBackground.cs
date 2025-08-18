//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 背景实现类
/// </summary>
internal class PowerPointBackground : IPowerPointBackground
{
    private readonly MsPowerPoint.ShapeRange _background;
    private bool _disposedValue;
    private IPowerPointFillFormat _fill;
    private IPowerPointColorScheme _colorScheme;

    /// <summary>
    /// 获取填充格式
    /// </summary>
    public IPowerPointFillFormat Fill
    {
        get
        {
            if (_fill == null && _background?.Fill != null)
            {
                _fill = new PowerPointFillFormat(_background.Fill);
            }
            return _fill;
        }
    }

    /// <summary>
    /// 获取颜色方案
    /// </summary>
    public IPowerPointColorScheme ColorScheme
    {
        get
        {
            if (_colorScheme == null && _background?.Parent is MsPowerPoint.Slide slide)
            {
                _colorScheme = new PowerPointColorScheme(slide.ColorScheme);
            }
            return _colorScheme;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _background?.Parent;

    /// <summary>
    /// 获取或设置背景样式
    /// </summary>
    public int Style
    {
        get => 0; // Background 没有直接的样式属性
        set { /* 不实现 */ }
    }

    /// <summary>
    /// 获取或设置背景类型
    /// </summary>
    public int Type
    {
        get => _background != null ? (int)_background.Type : 0;
    }

    /// <summary>
    /// 获取或设置是否显示背景图形
    /// </summary>
    public bool DisplayBackground
    {
        get => _background?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_background != null)
                _background.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="background">COM Background 对象</param>
    internal PowerPointBackground(MsPowerPoint.ShapeRange background)
    {
        _background = background;
        _disposedValue = false;
    }

    /// <summary>
    /// 应用纯色背景
    /// </summary>
    /// <param name="color">颜色</param>
    public void ApplySolidBackground(int color)
    {
        try
        {
            if (_background != null)
            {
                _background.Fill.Solid();
                _background.Fill.ForeColor.RGB = color;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply solid background.", ex);
        }
    }

    /// <summary>
    /// 应用渐变背景
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    /// <param name="color1">起始颜色</param>
    /// <param name="color2">结束颜色</param>
    public void ApplyGradientBackground(int style, int variant, int color1, int color2)
    {
        try
        {
            if (_background != null)
            {
                _background.Fill.TwoColorGradient((MsCore.MsoGradientStyle)style, variant);
                _background.Fill.ForeColor.RGB = color1;
                _background.Fill.BackColor.RGB = color2;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply gradient background.", ex);
        }
    }

    /// <summary>
    /// 应用图片背景
    /// </summary>
    /// <param name="pictureFile">图片文件路径</param>
    /// <param name="tile">是否平铺</param>
    public void ApplyPictureBackground(string pictureFile, bool tile = false)
    {
        if (string.IsNullOrEmpty(pictureFile))
            throw new ArgumentException("Picture file path cannot be null or empty.", nameof(pictureFile));

        if (!System.IO.File.Exists(pictureFile))
            throw new System.IO.FileNotFoundException("Picture file not found.", pictureFile);

        try
        {
            if (_background != null)
            {
                _background.Fill.UserPicture(pictureFile);
                // 平铺设置需要通过其他方式实现
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply picture background.", ex);
        }
    }

    /// <summary>
    /// 应用纹理背景
    /// </summary>
    /// <param name="textureType">纹理类型</param>
    public void ApplyTextureBackground(MsoPresetTexture textureType)
    {
        try
        {
            _background?.Fill.PresetTextured((MsCore.MsoPresetTexture)textureType);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply texture background.", ex);
        }
    }

    /// <summary>
    /// 应用主题背景
    /// </summary>
    /// <param name="themeIndex">主题索引</param>
    public void ApplyThemeBackground(int themeIndex)
    {
        try
        {
            // 主题背景应用需要通过幻灯片母版实现
            throw new NotImplementedException("Applying theme background is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply theme background.", ex);
        }
    }

    /// <summary>
    /// 重置背景
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_background != null)
            {
                _background.Fill.Solid();
                _background.Fill.ForeColor.RGB = 0xFFFFFF; // 白色
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset background.", ex);
        }
    }

    /// <summary>
    /// 应用到所有幻灯片
    /// </summary>
    public void ApplyToAll()
    {
        try
        {
            // 背景应用到所有幻灯片需要通过母版实现
            throw new NotImplementedException("Applying background to all slides is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply background to all slides.", ex);
        }
    }


    /// <summary>
    /// 获取背景信息
    /// </summary>
    /// <returns>背景信息字符串</returns>
    public string GetBackgroundInfo()
    {
        try
        {
            return $"Background - Type: {Type}, Visible: {DisplayBackground}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get background info.", ex);
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
            _fill?.Dispose();
            _colorScheme?.Dispose();
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

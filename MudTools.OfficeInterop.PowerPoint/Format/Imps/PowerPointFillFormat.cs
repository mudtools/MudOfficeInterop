//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 填充格式实现类
/// </summary>
internal class PowerPointFillFormat : IPowerPointFillFormat
{
    private readonly MsPowerPoint.FillFormat _fillFormat;
    private bool _disposedValue;
    private IPowerPointPictureFormat _picture;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _fillFormat?.Parent;

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    public int ForeColor
    {
        get => _fillFormat?.ForeColor.RGB ?? 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.ForeColor.RGB = value;
        }
    }

    /// <summary>
    /// 获取或设置背景色
    /// </summary>
    public int BackColor
    {
        get => _fillFormat?.BackColor.RGB ?? 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.BackColor.RGB = value;
        }
    }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _fillFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_fillFormat != null)
                _fillFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }


    /// <summary>
    /// 获取填充类型
    /// </summary>
    public int Type
    {
        get => _fillFormat != null ? (int)_fillFormat.Type : 0;
    }

    /// <summary>
    /// 获取或设置渐变颜色类型
    /// </summary>
    public int GradientColorType
    {
        get => _fillFormat != null ? (int)_fillFormat.GradientColorType : 0;
    }

    /// <summary>
    /// 获取或设置渐变样式
    /// </summary>
    public int GradientStyle
    {
        get => _fillFormat != null ? (int)_fillFormat.GradientStyle : 0;
    }

    /// <summary>
    /// 获取或设置渐变变体
    /// </summary>
    public int GradientVariant
    {
        get => _fillFormat?.GradientVariant ?? 0;
    }

    /// <summary>
    /// 获取或设置图案类型
    /// </summary>
    public int Pattern
    {
        get => _fillFormat != null ? (int)_fillFormat.Pattern : 0;
    }

    /// <summary>
    /// 获取或设置纹理类型
    /// </summary>
    public int TextureType
    {
        get => _fillFormat != null ? (int)_fillFormat.TextureType : 0;
        set
        {
            // TextureType 是只读属性
        }
    }

    /// <summary>
    /// 获取或设置纹理名称
    /// </summary>
    public string TextureName
    {
        get => _fillFormat?.TextureName ?? string.Empty;
        set
        {
            // TextureName 是只读属性
        }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="fillFormat">COM FillFormat 对象</param>
    internal PowerPointFillFormat(MsPowerPoint.FillFormat fillFormat)
    {
        _fillFormat = fillFormat; // 可以为 null
        _disposedValue = false;
    }

    /// <summary>
    /// 设置纯色填充
    /// </summary>
    public void Solid()
    {
        try
        {
            _fillFormat?.Solid();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set solid fill.", ex);
        }
    }

    /// <summary>
    /// 设置图案填充
    /// </summary>
    /// <param name="pattern">图案类型</param>
    public void Patterned(int pattern)
    {
        try
        {
            _fillFormat?.Patterned((MsCore.MsoPatternType)pattern);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set patterned fill.", ex);
        }
    }

    /// <summary>
    /// 设置渐变填充
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    /// <param name="presetGradientType">预设渐变类型</param>
    public void Gradient(int style, int variant, int presetGradientType)
    {
        try
        {
            _fillFormat?.OneColorGradient((MsCore.MsoGradientStyle)style, variant, presetGradientType);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set gradient fill.", ex);
        }
    }

    /// <summary>
    /// 设置纹理填充
    /// </summary>
    /// <param name="textureFile">纹理文件路径</param>
    /// <param name="textureType">纹理类型</param>
    public void Textured(string textureFile, MsoPresetTexture textureType = MsoPresetTexture.msoPresetTextureMixed)
    {
        if (string.IsNullOrEmpty(textureFile))
            throw new ArgumentException("Texture file path cannot be null or empty.", nameof(textureFile));

        try
        {
            if (System.IO.File.Exists(textureFile))
            {
                _fillFormat?.UserTextured(textureFile);
            }
            else
            {
                _fillFormat?.PresetTextured((MsCore.MsoPresetTexture)textureType);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set textured fill.", ex);
        }
    }

    /// <summary>
    /// 设置图片填充
    /// </summary>
    /// <param name="pictureFile">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    public void UserPicture(string pictureFile, bool linkToFile = false, bool saveWithDocument = true)
    {
        if (string.IsNullOrEmpty(pictureFile))
            throw new ArgumentException("Picture file path cannot be null or empty.", nameof(pictureFile));

        if (!System.IO.File.Exists(pictureFile))
            throw new System.IO.FileNotFoundException("Picture file not found.", pictureFile);

        try
        {
            _fillFormat?.UserPicture(pictureFile);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set picture fill.", ex);
        }
    }

    /// <summary>
    /// 设置预设纹理填充
    /// </summary>
    /// <param name="presetTexture">预设纹理类型</param>
    public void PresetTextured(MsoPresetTexture presetTexture)
    {
        try
        {
            _fillFormat?.PresetTextured((MsCore.MsoPresetTexture)presetTexture);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set preset textured fill.", ex);
        }
    }

    /// <summary>
    /// 设置预设渐变填充
    /// </summary>
    /// <param name="presetGradientType">预设渐变类型</param>
    public void PresetGradient(MsoPresetGradientType presetGradientType)
    {
        try
        {
            _fillFormat?.PresetGradient(MsCore.MsoGradientStyle.msoGradientHorizontal, 1, (MsCore.MsoPresetGradientType)presetGradientType);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set preset gradient fill.", ex);
        }
    }

    /// <summary>
    /// 设置自定义颜色渐变
    /// </summary>
    /// <param name="color1">起始颜色</param>
    /// <param name="color2">结束颜色</param>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    public void TwoColorGradient(int color1, int color2, int style, int variant)
    {
        try
        {
            _fillFormat?.TwoColorGradient((MsCore.MsoGradientStyle)style, variant);
            _fillFormat.ForeColor.RGB = color1;
            _fillFormat.BackColor.RGB = color2;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set two-color gradient fill.", ex);
        }
    }

    /// <summary>
    /// 重置填充格式
    /// </summary>
    public void Reset()
    {
        try
        {
            _fillFormat?.Solid();
            //_fillFormat?.ForeColor.RGB = 0; // 黑色
            //_fillFormat?.BackColor.RGB = 0xFFFFFF; // 白色
            //_fillFormat?.Visible = MsCore.MsoTriState.msoTrue;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset fill format.", ex);
        }
    }

    /// <summary>
    /// 复制填充格式
    /// </summary>
    /// <returns>复制的填充格式对象</returns>
    public IPowerPointFillFormat Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法
            throw new NotImplementedException("Duplicating fill format is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate fill format.", ex);
        }
    }

    /// <summary>
    /// 应用填充格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    public void ApplyTo(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        try
        {
            // 这需要具体的实现来应用填充格式到形状
            throw new NotImplementedException("Applying fill format to shape is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply fill format to shape.", ex);
        }
    }

    /// <summary>
    /// 设置填充颜色
    /// </summary>
    /// <param name="foregroundColor">前景色</param>
    /// <param name="backgroundColor">背景色</param>
    public void SetColors(int foregroundColor, int backgroundColor = 0)
    {
        try
        {
            if (_fillFormat != null)
            {
                _fillFormat.ForeColor.RGB = foregroundColor;
                if (backgroundColor != 0)
                    _fillFormat.BackColor.RGB = backgroundColor;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set fill colors.", ex);
        }
    }

    /// <summary>
    /// 获取填充信息
    /// </summary>
    /// <returns>填充信息字符串</returns>
    public string GetFillInfo()
    {
        try
        {
            return $"Fill Type: {Type}, ForeColor: {ForeColor}, BackColor: {BackColor}, Visible: {Visible}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get fill info.", ex);
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
            _picture?.Dispose();
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

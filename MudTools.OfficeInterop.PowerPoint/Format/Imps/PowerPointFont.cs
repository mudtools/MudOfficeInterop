//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 字体实现类
/// </summary>
internal class PowerPointFont : IPowerPointFont
{
    private readonly MsPowerPoint.Font _font;
    private bool _disposedValue;
    private IPowerPointFillFormat _fill;
    private IPowerPointLineFormat _line;
    private IPowerPointShadowFormat _shadowFormat;
    private IPowerPointSoftEdgeFormat _softEdge;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _font.Parent;

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    public string Name
    {
        get => _font.Name;
        set => _font.Name = value;
    }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    public float Size
    {
        get => _font.Size;
        set => _font.Size = value;
    }

    /// <summary>
    /// 获取或设置字体颜色
    /// </summary>
    public int Color
    {
        get => _font.Color.RGB;
        set => _font.Color.RGB = value;
    }

    /// <summary>
    /// 获取或设置是否加粗
    /// </summary>
    public bool Bold
    {
        get => _font.Bold == MsCore.MsoTriState.msoTrue;
        set => _font.Bold = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    public bool Italic
    {
        get => _font.Italic == MsCore.MsoTriState.msoTrue;
        set => _font.Italic = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取或设置下划线类型
    /// </summary>
    public int Underline
    {
        get => (int)_font.Underline;
        set => _font.Underline = (MsCore.MsoTriState)value;
    }

    /// <summary>
    /// 获取或设置上标
    /// </summary>
    public bool Subscript
    {
        get => _font.Subscript == MsCore.MsoTriState.msoTrue;
        set => _font.Subscript = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取或设置下标
    /// </summary>
    public bool Superscript
    {
        get => _font.Superscript == MsCore.MsoTriState.msoTrue;
        set => _font.Superscript = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }



    /// <summary>
    /// 获取或设置阴影效果
    /// </summary>
    public bool Shadow
    {
        get => _font.Shadow == MsCore.MsoTriState.msoTrue;
        set => _font.Shadow = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取或设置轮廓效果
    /// </summary>
    public bool Emboss
    {
        get => _font.Emboss == MsCore.MsoTriState.msoTrue;
        set => _font.Emboss = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }



    /// <summary>
    /// 获取或设置字体样式
    /// </summary>
    public int Style
    {
        get => 0; // PowerPoint Font 没有直接的 Style 属性
        set { /* 不实现 */ }
    }



    /// <summary>
    /// 获取或设置字体填充格式
    /// </summary>
    public IPowerPointFillFormat Fill
    {
        get
        {
            if (_fill == null)
            {
                _fill = new PowerPointFillFormat(null); // 需要具体的实现
            }
            return _fill;
        }
    }

    /// <summary>
    /// 获取或设置字体轮廓格式
    /// </summary>
    public IPowerPointLineFormat Line
    {
        get
        {
            if (_line == null)
            {
                _line = new PowerPointLineFormat(null); // 需要具体的实现
            }
            return _line;
        }
    }

    /// <summary>
    /// 获取或设置字体效果格式
    /// </summary>
    public IPowerPointShadowFormat ShadowFormat
    {
        get
        {
            if (_shadowFormat == null)
            {
                _shadowFormat = new PowerPointShadowFormat(null); // 需要具体的实现
            }
            return _shadowFormat;
        }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="font">COM Font 对象</param>
    internal PowerPointFont(MsPowerPoint.Font font)
    {
        _font = font ?? throw new ArgumentNullException(nameof(font));
        _disposedValue = false;
    }

    /// <summary>
    /// 复制字体设置
    /// </summary>
    /// <returns>复制的字体对象</returns>
    public IPowerPointFont Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法，返回新的实例
            throw new NotImplementedException("Duplicating font is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate font.", ex);
        }
    }

    /// <summary>
    /// 应用字体设置到指定文本范围
    /// </summary>
    /// <param name="textRange">目标文本范围</param>
    public void ApplyTo(IPowerPointTextRange textRange)
    {
        if (textRange == null)
            throw new ArgumentNullException(nameof(textRange));

        try
        {
            // 这需要具体的实现来应用字体到文本范围
            throw new NotImplementedException("Applying font to text range is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply font to text range.", ex);
        }
    }

    /// <summary>
    /// 重置字体设置为默认值
    /// </summary>
    public void Reset()
    {
        try
        {
            _font.Name = "Calibri";
            _font.Size = 18;
            _font.Color.RGB = 0; // 黑色
            _font.Bold = MsCore.MsoTriState.msoFalse;
            _font.Italic = MsCore.MsoTriState.msoFalse;
            _font.Underline = MsCore.MsoTriState.msoFalse;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset font.", ex);
        }
    }

    /// <summary>
    /// 设置字体基本属性
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="color">字体颜色</param>
    public void SetBasicProperties(string fontName = null, float fontSize = 0, int color = 0)
    {
        try
        {
            if (!string.IsNullOrEmpty(fontName))
                _font.Name = fontName;
            if (fontSize > 0)
                _font.Size = fontSize;
            if (color >= 0)
                _font.Color.RGB = color;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set basic font properties.", ex);
        }
    }

    /// <summary>
    /// 设置字体样式
    /// </summary>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="strikeThrough">删除线</param>
    public void SetStyle(bool bold = false, bool italic = false, int underline = 0, bool strikeThrough = false)
    {
        try
        {
            _font.Bold = bold ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            _font.Italic = italic ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            if (underline >= 0)
                _font.Underline = (MsCore.MsoTriState)underline;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set font style.", ex);
        }
    }

    /// <summary>
    /// 设置字体效果
    /// </summary>
    /// <param name="shadow">阴影效果</param>
    /// <param name="emboss">轮廓效果</param>
    /// <param name="imprint">浮雕效果</param>
    /// <param name="subscript">下标</param>
    /// <param name="superscript">上标</param>
    public void SetEffects(bool shadow = false, bool emboss = false, bool imprint = false, bool subscript = false, bool superscript = false)
    {
        try
        {
            _font.Shadow = shadow ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            _font.Emboss = emboss ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            _font.Subscript = subscript ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            _font.Superscript = superscript ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set font effects.", ex);
        }
    }

    /// <summary>
    /// 应用主题字体
    /// </summary>
    /// <param name="themeFontIndex">主题字体索引</param>
    public void ApplyThemeFont(int themeFontIndex = 1)
    {
        try
        {
            _font.Name = themeFontIndex == 1 ? "+mn-lt" : "+mj-lt"; // 主要字体或标题字体
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply theme font.", ex);
        }
    }

    /// <summary>
    /// 获取字体信息
    /// </summary>
    /// <returns>字体信息字符串</returns>
    public string GetFontInfo()
    {
        try
        {
            return $"Font: {Name}, Size: {Size}, Color: {Color}, Bold: {Bold}, Italic: {Italic}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get font info.", ex);
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
            _line?.Dispose();
            _shadowFormat?.Dispose();
            _softEdge?.Dispose();
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

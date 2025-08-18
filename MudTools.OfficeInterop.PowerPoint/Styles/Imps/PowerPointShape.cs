//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 形状实现类（精简版）
/// </summary>
internal class PowerPointShape : IPowerPointShape
{
    private readonly MsPowerPoint.Shape _shape;
    private bool _disposedValue;
    private IPowerPointTextFrame _textFrame;
    private IPowerPointFillFormat _fill;
    private IPowerPointLineFormat _line;
    private IPowerPointShadowFormat _shadow;
    private IPowerPointThreeDFormat _threeD;

    /// <summary>
    /// 获取或设置形状名称
    /// </summary>
    public string Name
    {
        get => _shape?.Name ?? string.Empty;
        set
        {
            if (_shape != null)
                _shape.Name = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取形状索引
    /// </summary>
    public int Index => _shape?.ZOrderPosition ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _shape?.Parent;

    /// <summary>
    /// 获取形状类型
    /// </summary>
    public MsoShapeType Type => _shape != null ? (MsoShapeType)_shape.Type : MsoShapeType.msoShapeTypeMixed;

    /// <summary>
    /// 获取或设置左边缘位置
    /// </summary>
    public double Left
    {
        get => _shape?.Left ?? 0;
        set
        {
            if (_shape != null)
                _shape.Left = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置上边缘位置
    /// </summary>
    public double Top
    {
        get => _shape?.Top ?? 0;
        set
        {
            if (_shape != null)
                _shape.Top = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置宽度
    /// </summary>
    public double Width
    {
        get => _shape?.Width ?? 0;
        set
        {
            if (_shape != null)
                _shape.Width = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置高度
    /// </summary>
    public double Height
    {
        get => _shape?.Height ?? 0;
        set
        {
            if (_shape != null)
                _shape.Height = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置旋转角度
    /// </summary>
    public double Rotation
    {
        get => _shape?.Rotation ?? 0;
        set
        {
            if (_shape != null)
                _shape.Rotation = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _shape?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shape != null)
                _shape.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置锁定纵横比
    /// </summary>
    public bool LockAspectRatio
    {
        get => _shape?.LockAspectRatio == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shape != null)
                _shape.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取Z轴顺序位置
    /// </summary>
    public int ZOrderPosition => _shape?.ZOrderPosition ?? 0;

    /// <summary>
    /// 获取文本框架
    /// </summary>
    public IPowerPointTextFrame TextFrame
    {
        get
        {
            if (_textFrame == null && _shape?.TextFrame != null)
            {
                _textFrame = new PowerPointTextFrame(_shape.TextFrame);
            }
            return _textFrame;
        }
    }

    /// <summary>
    /// 获取填充格式
    /// </summary>
    public IPowerPointFillFormat Fill
    {
        get
        {
            if (_fill == null && _shape?.Fill != null)
            {
                _fill = new PowerPointFillFormat(_shape.Fill);
            }
            return _fill;
        }
    }

    /// <summary>
    /// 获取线条格式
    /// </summary>
    public IPowerPointLineFormat Line
    {
        get
        {
            if (_line == null && _shape?.Line != null)
            {
                _line = new PowerPointLineFormat(_shape.Line);
            }
            return _line;
        }
    }

    /// <summary>
    /// 获取阴影格式
    /// </summary>
    public IPowerPointShadowFormat Shadow
    {
        get
        {
            if (_shadow == null && _shape?.Shadow != null)
            {
                _shadow = new PowerPointShadowFormat(_shape.Shadow);
            }
            return _shadow;
        }
    }

    /// <summary>
    /// 获取三维格式
    /// </summary>
    public IPowerPointThreeDFormat ThreeD
    {
        get
        {
            if (_threeD == null && _shape?.ThreeD != null)
            {
                _threeD = new PowerPointThreeDFormat(_shape.ThreeD);
            }
            return _threeD;
        }
    }

    /// <summary>
    /// 获取是否具有文本框架
    /// </summary>
    public bool HasTextFrame => _shape?.HasTextFrame == MsCore.MsoTriState.msoTrue;

    /// <summary>
    /// 获取形状ID
    /// </summary>
    public int ID => _shape?.Id ?? 0;

    /// <summary>
    /// 获取或设置替代文本
    /// </summary>
    public string AlternativeText
    {
        get => _shape?.AlternativeText ?? string.Empty;
        set
        {
            if (_shape != null)
                _shape.AlternativeText = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取是否为组形状
    /// </summary>
    public bool IsGroup => _shape?.Type == MsCore.MsoShapeType.msoGroup;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="shape">COM Shape 对象</param>
    internal PowerPointShape(MsPowerPoint.Shape shape)
    {
        _shape = shape ?? throw new ArgumentNullException(nameof(shape));
        _disposedValue = false;
    }

    /// <summary>
    /// 选择形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        try
        {
            _shape?.Select(replace ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select shape.", ex);
        }
    }

    /// <summary>
    /// 复制形状
    /// </summary>
    public void Copy()
    {
        try
        {
            _shape?.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy shape.", ex);
        }
    }

    /// <summary>
    /// 剪切形状
    /// </summary>
    public void Cut()
    {
        try
        {
            _shape?.Cut();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to cut shape.", ex);
        }
    }

    /// <summary>
    /// 删除形状
    /// </summary>
    public void Delete()
    {
        try
        {
            _shape?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete shape.", ex);
        }
    }

    /// <summary>
    /// 移动形状
    /// </summary>
    /// <param name="x">水平移动距离</param>
    /// <param name="y">垂直移动距离</param>
    public void Move(double x, double y)
    {
        try
        {
            if (_shape != null)
            {
                _shape.Left += (float)x;
                _shape.Top += (float)y;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to move shape.", ex);
        }
    }

    /// <summary>
    /// 调整形状大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    public void Scale(double width, double height)
    {
        try
        {
            if (_shape != null)
            {
                _shape.Width = (float)width;
                _shape.Height = (float)height;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to scale shape.", ex);
        }
    }

    /// <summary>
    /// 旋转形状
    /// </summary>
    /// <param name="angle">旋转角度</param>
    public void Rotate(double angle)
    {
        try
        {
            if (_shape != null)
            {
                _shape.Rotation += (float)angle;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to rotate shape.", ex);
        }
    }

    /// <summary>
    /// 水平翻转
    /// </summary>
    public void FlipHorizontal()
    {
        try
        {
            _shape?.Flip(MsCore.MsoFlipCmd.msoFlipHorizontal);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to flip shape horizontally.", ex);
        }
    }

    /// <summary>
    /// 垂直翻转
    /// </summary>
    public void FlipVertical()
    {
        try
        {
            _shape?.Flip(MsCore.MsoFlipCmd.msoFlipVertical);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to flip shape vertically.", ex);
        }
    }

    /// <summary>
    /// 设置Z轴顺序
    /// </summary>
    /// <param name="position">位置</param>
    public void ZOrder(int position)
    {
        try
        {
            _shape?.ZOrder((MsCore.MsoZOrderCmd)position);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set Z-order.", ex);
        }
    }

    /// <summary>
    /// 取消组合
    /// </summary>
    /// <returns>形状范围</returns>
    public IPowerPointShapeRange Ungroup()
    {
        try
        {
            if (_shape.Type == MsCore.MsoShapeType.msoGroup)
            {
                var shapeRange = _shape.Ungroup();
                return new PowerPointShapeRange(shapeRange);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to ungroup shape.", ex);
        }
    }

    /// <summary>
    /// 获取文本内容
    /// </summary>
    /// <returns>文本内容</returns>
    public string GetText()
    {
        try
        {
            return _shape?.TextFrame?.TextRange?.Text ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get text from shape.", ex);
        }
    }

    /// <summary>
    /// 设置文本内容
    /// </summary>
    /// <param name="text">文本内容</param>
    public void SetText(string text)
    {
        try
        {
            if (_shape?.TextFrame?.TextRange != null)
            {
                _shape.TextFrame.TextRange.Text = text ?? string.Empty;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set text for shape.", ex);
        }
    }

    /// <summary>
    /// 替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    public int ReplaceText(string findText, string replaceText)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        try
        {
            var r = _shape?.TextFrame?.TextRange?.Replace(findText, replaceText ?? string.Empty);
            return r != null ? r.Count : 0;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in shape.", ex);
        }
    }

    /// <summary>
    /// 设置填充颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    public void SetFillColor(int color)
    {
        try
        {
            if (_shape?.Fill != null)
            {
                _shape.Fill.ForeColor.RGB = color;
                _shape.Fill.Visible = MsCore.MsoTriState.msoTrue;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set fill color.", ex);
        }
    }

    /// <summary>
    /// 设置线条颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    public void SetLineColor(int color)
    {
        try
        {
            if (_shape?.Line != null)
            {
                _shape.Line.ForeColor.RGB = color;
                _shape.Line.Visible = MsCore.MsoTriState.msoTrue;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line color.", ex);
        }
    }

    /// <summary>
    /// 设置线条粗细
    /// </summary>
    /// <param name="weight">粗细</param>
    public void SetLineWeight(float weight)
    {
        try
        {
            if (_shape?.Line != null)
            {
                _shape.Line.Weight = weight;
                _shape.Line.Visible = MsCore.MsoTriState.msoTrue;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line weight.", ex);
        }
    }

    /// <summary>
    /// 应用阴影效果
    /// </summary>
    /// <param name="shadowType">阴影类型</param>
    public void ApplyShadow(int shadowType = 1)
    {
        try
        {
            if (_shape?.Shadow != null)
            {
                _shape.Shadow.Visible = MsCore.MsoTriState.msoTrue;
                _shape.Shadow.Type = (MsCore.MsoShadowType)shadowType;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply shadow effect.", ex);
        }
    }

    /// <summary>
    /// 应用三维效果
    /// </summary>
    /// <param name="depth">深度</param>
    public void Apply3DEffect(float depth = 10)
    {
        try
        {
            if (_shape?.ThreeD != null)
            {
                _shape.ThreeD.Visible = MsCore.MsoTriState.msoTrue;
                _shape.ThreeD.Depth = depth;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply 3D effect.", ex);
        }
    }

    /// <summary>
    /// 导出为图片
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="filterName">格式</param>
    public void Export(string fileName, int filterName = 2)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            _shape?.Export(fileName, (MsPowerPoint.PpShapeFormat)filterName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export shape to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 刷新显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            _shape?.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh shape.", ex);
        }
    }

    /// <summary>
    /// 设置透明度
    /// </summary>
    /// <param name="transparency">透明度</param>
    public void SetTransparency(float transparency)
    {
        try
        {
            if (_shape?.Fill != null)
            {
                _shape.Fill.Transparency = transparency;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set transparency.", ex);
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
            _textFrame?.Dispose();
            _fill?.Dispose();
            _line?.Dispose();
            _shadow?.Dispose();
            _threeD?.Dispose();
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

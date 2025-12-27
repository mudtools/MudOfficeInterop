//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 线条格式实现类
/// </summary>
internal class PowerPointLineFormat : IPowerPointLineFormat
{
    private readonly MsPowerPoint.LineFormat _lineFormat;
    private bool _disposedValue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _lineFormat?.Parent;

    /// <summary>
    /// 获取或设置线条样式
    /// </summary>
    public int Style
    {
        get => _lineFormat != null ? (int)_lineFormat.Style : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Style = (MsCore.MsoLineStyle)value;
        }
    }

    /// <summary>
    /// 获取或设置线条粗细
    /// </summary>
    public float Weight
    {
        get => _lineFormat?.Weight ?? 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Weight = value;
        }
    }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    public int ForeColor
    {
        get => _lineFormat?.ForeColor.RGB ?? 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.ForeColor.RGB = value;
        }
    }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _lineFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_lineFormat != null)
                _lineFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }



    /// <summary>
    /// 获取或设置虚线样式
    /// </summary>
    public int DashStyle
    {
        get => _lineFormat != null ? (int)_lineFormat.DashStyle : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.DashStyle = (MsCore.MsoLineDashStyle)value;
        }
    }


    /// <summary>
    /// 获取或设置起始箭头样式
    /// </summary>
    public int BeginArrowheadStyle
    {
        get => _lineFormat != null ? (int)_lineFormat.BeginArrowheadStyle : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.BeginArrowheadStyle = (MsCore.MsoArrowheadStyle)value;
        }
    }

    /// <summary>
    /// 获取或设置结束箭头样式
    /// </summary>
    public int EndArrowheadStyle
    {
        get => _lineFormat != null ? (int)_lineFormat.EndArrowheadStyle : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.EndArrowheadStyle = (MsCore.MsoArrowheadStyle)value;
        }
    }

    /// <summary>
    /// 获取或设置起始箭头宽度
    /// </summary>
    public int BeginArrowheadWidth
    {
        get => _lineFormat != null ? (int)_lineFormat.BeginArrowheadWidth : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.BeginArrowheadWidth = (MsCore.MsoArrowheadWidth)value;
        }
    }

    /// <summary>
    /// 获取或设置结束箭头宽度
    /// </summary>
    public int EndArrowheadWidth
    {
        get => _lineFormat != null ? (int)_lineFormat.EndArrowheadWidth : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.EndArrowheadWidth = (MsCore.MsoArrowheadWidth)value;
        }
    }

    /// <summary>
    /// 获取或设置起始箭头长度
    /// </summary>
    public int BeginArrowheadLength
    {
        get => _lineFormat != null ? (int)_lineFormat.BeginArrowheadLength : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.BeginArrowheadLength = (MsCore.MsoArrowheadLength)value;
        }
    }

    /// <summary>
    /// 获取或设置结束箭头长度
    /// </summary>
    public int EndArrowheadLength
    {
        get => _lineFormat != null ? (int)_lineFormat.EndArrowheadLength : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.EndArrowheadLength = (MsCore.MsoArrowheadLength)value;
        }
    }


    /// <summary>
    /// 获取线条类型
    /// </summary>
    public int Type
    {
        get => _lineFormat != null ? (int)_lineFormat.Style : 0;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="lineFormat">COM LineFormat 对象</param>
    internal PowerPointLineFormat(MsPowerPoint.LineFormat lineFormat)
    {
        _lineFormat = lineFormat; // 可以为 null
        _disposedValue = false;
    }

    /// <summary>
    /// 设置箭头样式
    /// </summary>
    /// <param name="beginStyle">起始箭头样式</param>
    /// <param name="endStyle">结束箭头样式</param>
    /// <param name="beginWidth">起始箭头宽度</param>
    /// <param name="endWidth">结束箭头宽度</param>
    /// <param name="beginLength">起始箭头长度</param>
    /// <param name="endLength">结束箭头长度</param>
    public void SetArrowheads(int beginStyle = 0, int endStyle = 0, int beginWidth = 1, int endWidth = 1, int beginLength = 1, int endLength = 1)
    {
        try
        {
            if (_lineFormat != null)
            {
                _lineFormat.BeginArrowheadStyle = (MsCore.MsoArrowheadStyle)beginStyle;
                _lineFormat.EndArrowheadStyle = (MsCore.MsoArrowheadStyle)endStyle;
                _lineFormat.BeginArrowheadWidth = (MsCore.MsoArrowheadWidth)beginWidth;
                _lineFormat.EndArrowheadWidth = (MsCore.MsoArrowheadWidth)endWidth;
                _lineFormat.BeginArrowheadLength = (MsCore.MsoArrowheadLength)beginLength;
                _lineFormat.EndArrowheadLength = (MsCore.MsoArrowheadLength)endLength;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set arrowheads.", ex);
        }
    }


    /// <summary>
    /// 重置线条格式
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_lineFormat != null)
            {
                _lineFormat.Style = MsCore.MsoLineStyle.msoLineSingle;
                _lineFormat.Weight = 1;
                _lineFormat.ForeColor.RGB = 0; // 黑色
                _lineFormat.Visible = MsCore.MsoTriState.msoTrue;
                _lineFormat.DashStyle = MsCore.MsoLineDashStyle.msoLineSolid;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset line format.", ex);
        }
    }

    /// <summary>
    /// 复制线条格式
    /// </summary>
    /// <returns>复制的线条格式对象</returns>
    public IPowerPointLineFormat Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法
            throw new NotImplementedException("Duplicating line format is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate line format.", ex);
        }
    }

    /// <summary>
    /// 应用线条格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    public void ApplyTo(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        try
        {
            // 这需要具体的实现来应用线条格式到形状
            throw new NotImplementedException("Applying line format to shape is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply line format to shape.", ex);
        }
    }

    /// <summary>
    /// 设置线条粗细
    /// </summary>
    /// <param name="weight">线条粗细</param>
    public void SetWeight(float weight)
    {
        try
        {
            if (_lineFormat != null) _lineFormat.Weight = weight;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line weight.", ex);
        }
    }

    /// <summary>
    /// 设置线条颜色
    /// </summary>
    /// <param name="color">线条颜色</param>
    public void SetColor(int color)
    {
        try
        {
            if (_lineFormat != null) _lineFormat.ForeColor.RGB = color;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line color.", ex);
        }
    }

    /// <summary>
    /// 设置线条样式
    /// </summary>
    /// <param name="style">线条样式</param>
    public void SetStyle(int style)
    {
        try
        {
            if (_lineFormat != null) _lineFormat.Style = (MsCore.MsoLineStyle)style;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line style.", ex);
        }
    }

    /// <summary>
    /// 获取线条信息
    /// </summary>
    /// <returns>线条信息字符串</returns>
    public string GetLineInfo()
    {
        try
        {
            return $"Line Style: {Style}, Weight: {Weight}, Color: {ForeColor}, Visible: {Visible}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get line info.", ex);
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

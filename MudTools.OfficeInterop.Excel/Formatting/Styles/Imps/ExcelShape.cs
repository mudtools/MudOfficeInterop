//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Shape 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Shape 对象的安全访问和资源管理
/// </summary>
internal class ExcelShape : IExcelShape
{
    /// <summary>
    /// 底层的 COM Shape 对象
    /// </summary>
    private MsExcel.Shape _shape;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelShape 实例
    /// </summary>
    /// <param name="shape">底层的 COM Shape 对象</param>
    internal ExcelShape(MsExcel.Shape? shape)
    {
        _shape = shape;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放子COM组件
                (_fill as ExcelFillFormat)?.Dispose();
                (_line as ExcelLineFormat)?.Dispose();
                (_textFrame as ExcelTextFrame)?.Dispose();
                (_shadow as ExcelShadowFormat)?.Dispose();
                (_threeD as ExcelThreeDFormat)?.Dispose();
                (_topLeftCell as ExcelRange)?.Dispose();
                (_bottomRightCell as ExcelRange)?.Dispose();
                (_chart as ExcelChart)?.Dispose();

                // 释放底层COM对象
                if (_shape != null)
                    Marshal.ReleaseComObject(_shape);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _shape = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置形状的名称
    /// </summary>
    public string Name
    {
        get => _shape?.Name?.ToString();
        set
        {
            if (_shape != null && value != null)
                _shape.Name = value;
        }
    }

    public IExcelOLEFormat OLEFormat
    {
        get => new ExcelOLEFormat(_shape.OLEFormat);
    }

    public IExcelGroupShapes GroupItems
    {
        get => new ExcelGroupShapes(_shape.GroupItems);
    }

    /// <summary>
    /// 获取形状的类型
    /// </summary>
    public MsoShapeType Type => _shape != null ? (MsoShapeType)_shape.Type : MsoShapeType.msoShapeTypeMixed;

    /// <summary>
    /// 获取形状的ID
    /// </summary>
    public int ID => _shape != null ? _shape.ID : 0;

    public bool LockAspectRatio
    {
        get => _shape.LockAspectRatio == MsCore.MsoTriState.msoTrue;
        set => _shape.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取形状的父对象
    /// </summary>
    public object Parent => _shape?.Parent;

    public XlPlacement Placement
    {
        get
        {
            return (XlPlacement)_shape?.Placement;
        }
        set
        {
            _shape.Placement = (MsExcel.XlPlacement)value;
        }
    }

    #endregion

    #region 位置和大小

    /// <summary>
    /// 获取或设置形状的左边距
    /// </summary>
    public float Left
    {
        get => _shape != null ? _shape.Left : 0;
        set
        {
            if (_shape != null)
                _shape.Left = value;
        }
    }

    /// <summary>
    /// 获取或设置形状的顶边距
    /// </summary>
    public float Top
    {
        get => _shape != null ? _shape.Top : 0;
        set
        {
            if (_shape != null)
                _shape.Top = value;
        }
    }

    /// <summary>
    /// 获取或设置形状的宽度
    /// </summary>
    public float Width
    {
        get => _shape != null ? _shape.Width : 0;
        set
        {
            if (_shape != null)
                _shape.Width = value;
        }
    }

    /// <summary>
    /// 获取或设置形状的高度
    /// </summary>
    public float Height
    {
        get => _shape != null ? _shape.Height : 0;
        set
        {
            if (_shape != null)
                _shape.Height = value;
        }
    }

    /// <summary>
    /// 获取或设置形状的旋转角度
    /// </summary>
    public float Rotation
    {
        get => _shape != null ? _shape.Rotation : 0;
        set
        {
            if (_shape != null)
                _shape.Rotation = value;
        }
    }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状是否可见
    /// </summary>
    public bool Visible
    {
        get => _shape != null && Convert.ToBoolean(_shape.Visible);
        set
        {
            if (_shape != null)
                _shape.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse; ;
        }
    }

    /// <summary>
    /// 获取或设置形状是否锁定
    /// </summary>
    public bool Locked
    {
        get => _shape != null && Convert.ToBoolean(_shape.Locked);
        set
        {
            if (_shape != null)
                _shape.Locked = value;
        }
    }



    #endregion

    #region 格式设置

    /// <summary>
    /// 填充格式对象缓存
    /// </summary>
    private IExcelFillFormat _fill;

    /// <summary>
    /// 获取形状的填充格式对象
    /// </summary>
    public IExcelFillFormat Fill => _fill ?? (_fill = new ExcelFillFormat(_shape?.Fill));

    /// <summary>
    /// 线条格式对象缓存
    /// </summary>
    private IExcelLineFormat _line;

    /// <summary>
    /// 获取形状的线条格式对象
    /// </summary>
    public IExcelLineFormat Line => _line ??= new ExcelLineFormat(_shape?.Line);

    /// <summary>
    /// 文本框架对象缓存
    /// </summary>
    private IExcelTextFrame _textFrame;

    /// <summary>
    /// 获取形状的文本框架对象
    /// </summary>
    public IExcelTextFrame TextFrame => _textFrame ??= new ExcelTextFrame(_shape?.TextFrame);

    /// <summary>
    /// 阴影格式对象缓存
    /// </summary>
    private IExcelShadowFormat _shadow;

    /// <summary>
    /// 获取形状的阴影格式对象
    /// </summary>
    public IExcelShadowFormat Shadow => _shadow ??= new ExcelShadowFormat(_shape?.Shadow);

    /// <summary>
    /// 三维格式对象缓存
    /// </summary>
    private IExcelThreeDFormat _threeD;

    /// <summary>
    /// 获取形状的三维格式对象
    /// </summary>
    public IExcelThreeDFormat ThreeD => _threeD ??= new ExcelThreeDFormat(_shape?.ThreeD);

    #endregion

    #region 文本属性

    /// <summary>
    /// 获取或设置形状中的文本内容
    /// </summary>
    public string Text
    {
        get => _shape?.TextFrame?.Characters()?.Text?.ToString();
        set
        {
            if (_shape?.TextFrame?.Characters() != null && value != null)
                _shape.TextFrame.Characters().Text = value;
        }
    }

    /// <summary>
    /// 获取或设置形状中文本的自动调整大小
    /// </summary>
    public bool AutoSize
    {
        get => _shape?.TextFrame != null && Convert.ToBoolean(_shape.TextFrame.AutoSize);
        set
        {
            if (_shape?.TextFrame != null)
                _shape.TextFrame.AutoSize = value;
        }
    }

    /// <summary>
    /// 获取或设置形状中文本的水平对齐方式
    /// </summary>
    public XlHAlign HorizontalAlignment
    {
        get => _shape?.TextFrame != null ? (XlHAlign)_shape.TextFrame.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_shape?.TextFrame != null)
                _shape.TextFrame.HorizontalAlignment = (MsExcel.XlHAlign)value;
        }
    }

    /// <summary>
    /// 获取或设置形状中文本的垂直对齐方式
    /// </summary>
    public XlVAlign VerticalAlignment
    {
        get => _shape?.TextFrame != null ? (XlVAlign)_shape.TextFrame.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_shape?.TextFrame != null)
                _shape.TextFrame.VerticalAlignment = (MsExcel.XlVAlign)value;
        }
    }

    #endregion

    #region 操作方法



    /// <summary>
    /// 选择形状
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    public void Select(bool replace = true)
    {
        _shape?.Select(replace);
    }

    /// <summary>
    /// 复制形状
    /// </summary>
    public void Copy()
    {
        _shape?.Copy();
    }

    public void CopyPicture(XlPictureAppearance? Appearance, XlCopyPictureFormat? Format)
    {
        _shape?.CopyPicture(Appearance.ComArgsVal(), Format.ComArgsVal());
    }

    /// <summary>
    /// 剪切形状
    /// </summary>
    public void Cut()
    {
        _shape?.Cut();
    }

    /// <summary>
    /// 删除形状
    /// </summary>
    public void Delete()
    {
        _shape?.Delete();
    }

    public void ScaleHeight(float Factor, bool RelativeToOriginalSize, double Scale)
    {
        _shape?.ScaleHeight(Factor, RelativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse, Scale);
    }

    public void ScaleWidth(float Factor, bool RelativeToOriginalSize, double Scale)
    {
        _shape?.ScaleWidth(Factor, RelativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse, Scale);
    }


    /// <summary>
    /// 调整形状大小
    /// </summary>
    /// <param name="widthScale">宽度缩放比例</param>
    /// <param name="heightScale">高度缩放比例</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小</param>
    public void Scale(double widthScale, double heightScale, bool relativeToOriginalSize = false)
    {
        if (_shape != null)
        {
            _shape.ScaleWidth((float)widthScale,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                MsExcel.XlScaleType.xlScaleLinear);
            _shape.ScaleHeight((float)heightScale,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                MsExcel.XlScaleType.xlScaleLinear);
        }
    }

    /// <summary>
    /// 移动形状
    /// </summary>
    /// <param name="leftIncrement">左边距增量</param>
    /// <param name="topIncrement">顶边距增量</param>
    public void Move(double leftIncrement, double topIncrement)
    {
        _shape?.IncrementLeft((float)leftIncrement);
        _shape?.IncrementTop((float)topIncrement);
    }

    /// <summary>
    /// 旋转形状
    /// </summary>
    /// <param name="rotationIncrement">旋转角度增量（度）</param>
    public void Rotate(double rotationIncrement)
    {
        _shape?.IncrementRotation((float)rotationIncrement);
    }

    /// <summary>
    /// 将形状置于最前面
    /// </summary>
    public void BringToFront()
    {
        _shape?.ZOrder(MsCore.MsoZOrderCmd.msoBringToFront);
    }

    /// <summary>
    /// 将形状置于最后面
    /// </summary>
    public void SendToBack()
    {
        _shape?.ZOrder(MsCore.MsoZOrderCmd.msoSendToBack);
    }

    /// <summary>
    /// 取消组合形状
    /// </summary>
    /// <returns>取消组合后的形状集合</returns>
    public IExcelShapes Ungroup()
    {
        var shapes = _shape?.Ungroup() as MsExcel.Shapes;
        return shapes != null ? new ExcelShapes(shapes) : null;
    }

    /// <summary>
    /// 应用自动调整选项
    /// </summary>
    public void Apply()
    {
        _shape?.Apply();
    }

    /// <summary>
    /// 复制形状的格式
    /// </summary>
    public void PickUp()
    {
        _shape?.PickUp();
    }


    #endregion

    #region 层次结构

    /// <summary>
    /// 左上角单元格缓存
    /// </summary>
    private IExcelRange _topLeftCell;

    /// <summary>
    /// 获取形状所在的区域对象（左上角）
    /// </summary>
    public IExcelRange TopLeftCell => _topLeftCell ??= new ExcelRange(_shape?.TopLeftCell);

    /// <summary>
    /// 右下角单元格缓存
    /// </summary>
    private IExcelRange _bottomRightCell;

    /// <summary>
    /// 获取形状所在的区域对象（右下角）
    /// </summary>
    public IExcelRange BottomRightCell => _bottomRightCell ??= new ExcelRange(_shape?.BottomRightCell);

    /// <summary>
    /// 图表对象缓存
    /// </summary>
    private IExcelChart _chart;

    /// <summary>
    /// 获取形状所在的图表对象（如果是图表）
    /// </summary>
    public IExcelChart Chart => _chart ??= new ExcelChart(_shape?.Chart);

    #endregion
}
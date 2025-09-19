
namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelShapeNode
// =============================================
internal class ExcelShapeNode : IExcelShapeNode
{
    internal MsExcel.ShapeNode _shapeNode;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="shapeNode">原始的 COM ShapeNode 对象</param>
    internal ExcelShapeNode(MsExcel.ShapeNode shapeNode)
    {
        _shapeNode = shapeNode ?? throw new ArgumentNullException(nameof(shapeNode));
    }

    /// <summary>
    /// 获取此节点所属的父对象（通常是 ShapeNodes 集合）。
    /// </summary>
    public object Parent => _shapeNode.Parent;

    /// <summary>
    /// 获取此节点所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_shapeNode.Application as MsExcel.Application);


    /// <summary>
    /// 获取节点类型（角点或曲线点）。
    /// </summary>
    public MsoEditingType EditingType => _shapeNode.EditingType.EnumConvert(MsoEditingType.msoEditingAuto);

    public MsoSegmentType SegmentType
    {
        get
        {
            var segmentType = _shapeNode.SegmentType;
            return segmentType.EnumConvert(MsoSegmentType.msoSegmentLine);
        }
    }

    /// <summary>
    /// 获取节点在工作表中的 X 坐标（单位：磅，points）。
    /// </summary>
    public float PointsX
    {
        get
        {
            var points = GetPointsArray();
            return points?.GetValue(1, 1) is float x ? x : 0f;
        }
    }

    /// <summary>
    /// 获取节点在工作表中的 Y 坐标（单位：磅，points）。
    /// </summary>
    public float PointsY
    {
        get
        {
            var points = GetPointsArray();
            return points?.GetValue(1, 2) is float y ? y : 0f;
        }
    }

    /// <summary>
    /// 获取第一个控制点相对于节点的 X 偏移量（仅对曲线节点有效）。
    /// </summary>
    public float? ControlPoint1X
    {
        get
        {
            var points = GetPointsArray();
            if (points == null || points.GetLength(0) < 3) return null;
            return points.GetValue(2, 1) as float?;
        }
    }

    /// <summary>
    /// 获取第一个控制点相对于节点的 Y 偏移量（仅对曲线节点有效）。
    /// </summary>
    public float? ControlPoint1Y
    {
        get
        {
            var points = GetPointsArray();
            if (points == null || points.GetLength(0) < 3) return null;
            return points.GetValue(2, 2) as float?;
        }
    }

    /// <summary>
    /// 获取第二个控制点相对于节点的 X 偏移量（仅对曲线节点有效）。
    /// </summary>
    public float? ControlPoint2X
    {
        get
        {
            var points = GetPointsArray();
            if (points == null || points.GetLength(0) < 3) return null;
            return points.GetValue(3, 1) as float?;
        }
    }

    /// <summary>
    /// 获取第二个控制点相对于节点的 Y 偏移量（仅对曲线节点有效）。
    /// </summary>
    public float? ControlPoint2Y
    {
        get
        {
            var points = GetPointsArray();
            if (points == null || points.GetLength(0) < 3) return null;
            return points.GetValue(3, 2) as float?;
        }
    }

    /// <summary>
    /// 内部辅助方法：获取节点的 Points 数组（二维 object 数组）。
    /// </summary>
    private object[,] GetPointsArray()
    {
        try
        {
            return _shapeNode.Points as object[,];
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"获取 ShapeNode.Points 失败: {ex.Message}");
            return null;
        }
    }

    #region IDisposable Support

    /// <summary>
    /// 释放托管和非托管资源。
    /// </summary>
    /// <param name="disposing">是否正在释放托管资源</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_shapeNode != null)
            {
                Marshal.ReleaseComObject(_shapeNode);
                _shapeNode = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelShapeNode()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
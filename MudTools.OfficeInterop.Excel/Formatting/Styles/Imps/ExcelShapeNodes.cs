
namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelShapeNodes
// =============================================
internal class ExcelShapeNodes : IExcelShapeNodes
{
    internal MsExcel.ShapeNodes _shapeNodes;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelPictures));
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="shapeNodes">原始的 COM ShapeNodes 对象</param>
    internal ExcelShapeNodes(MsExcel.ShapeNodes shapeNodes)
    {
        _shapeNodes = shapeNodes ?? throw new ArgumentNullException(nameof(shapeNodes));
    }

    /// <summary>
    /// 获取集合中节点的总数。
    /// </summary>
    public int Count => _shapeNodes.Count;

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的节点。
    /// </summary>
    /// <param name="index">节点索引（1-based）</param>
    /// <returns>对应的节点对象</returns>
    public IExcelShapeNode this[int index]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelShapeNodes));
            try
            {
                return new ExcelShapeNode((MsExcel.ShapeNode)_shapeNodes.Item(index));
            }
            catch (Exception ex)
            {
                log.Error($"获取第 {index} 个形状节点失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Shape）。
    /// </summary>
    public object Parent => _shapeNodes.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_shapeNodes.Application as MsExcel.Application);

    /// <summary>
    /// 在指定索引位置插入一个新节点。
    /// </summary>
    /// <param name="index">插入位置（从 1 开始）</param>
    /// <param name="segmentType">节点类型（直线或曲线）</param>
    /// <param name="editingType">节点类型（角点或曲线）</param>
    /// <param name="x1">节点 X 坐标</param>
    /// <param name="y1">节点 Y 坐标</param>
    /// <param name="x2">第一个控制点 X 偏移（仅曲线节点需要）</param>
    /// <param name="y2">第一个控制点 Y 偏移（仅曲线节点需要）</param>
    /// <param name="x3">第二个控制点 X 偏移（仅曲线节点需要）</param>
    /// <param name="y3">第二个控制点 Y 偏移（仅曲线节点需要）</param>
    /// <returns>新创建的节点对象</returns>
    public void Insert(
        int index,
        MsoSegmentType segmentType,
        MsoEditingType editingType,
        float x1, float y1,
        float x2 = 0, float y2 = 0,
        float x3 = 0, float y3 = 0)
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelShapeNodes));

        try
        {
            _shapeNodes.Insert(index,
            segmentType.EnumConvert(MsCore.MsoSegmentType.msoSegmentLine),
            editingType.EnumConvert(MsCore.MsoEditingType.msoEditingCorner),
             x1, y1,
             x2, y2, x3, y3);
        }
        catch (Exception ex)
        {
            log.Error($"在位置 {index} 插入形状节点失败: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// 设置指定索引节点的属性。
    /// </summary>
    /// <param name="index">节点索引（1-based）</param>
    /// <param name="x1">节点 X 坐标</param>
    /// <param name="y1">节点 Y 坐标</param>
    public void SetPosition(int index, float x1, float y1)
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelShapeNodes));

        try
        {
            _shapeNodes.SetPosition(index, x1, y1);
        }
        catch (Exception ex)
        {
            log.Error($"设置第 {index} 个形状节点位置失败: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// 内部辅助方法：根据参数构建 Points 数组。
    /// </summary>
    private object[,] BuildPointsArray(
        MsoEditingType editingType,
        float x1, float y1,
        float? x2, float? y2,
        float? x3, float? y3)
    {
        int rows = editingType == MsoEditingType.msoEditingCorner ? 1 : 3;
        var points = new object[rows, 2];

        // 节点主坐标
        points[0, 0] = x1;
        points[0, 1] = y1;

        if (editingType != MsoEditingType.msoEditingCorner)
        {
            // 控制点1
            points[1, 0] = x2 ?? 0f;
            points[1, 1] = y2 ?? 0f;

            // 控制点2
            points[2, 0] = x3 ?? 0f;
            points[2, 1] = y3 ?? 0f;
        }

        return points;
    }

    #region IEnumerable<IExcelShapeNode> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelShapeNode> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelShapeNodes));

        for (int i = 1; i <= _shapeNodes.Count; i++)
        {
            yield return new ExcelShapeNode(_shapeNodes.Item(i));
        }
    }

    /// <summary>
    /// 非泛型枚举器支持。
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

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
            try
            {
                if (_shapeNodes != null)
                {
                    Marshal.ReleaseComObject(_shapeNodes);
                    _shapeNodes = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"释放 ShapeNodes 时发生异常: {ex.Message}");
                // 忽略释放异常，避免掩盖更严重的问题
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelShapeNodes()
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
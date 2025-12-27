//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 形状集合实现类
/// </summary>
internal class PowerPointShapes : IPowerPointShapes
{
    private readonly MsPowerPoint.Shapes _shapes;
    private bool _disposedValue;

    /// <summary>
    /// 获取形状数量
    /// </summary>
    public int Count => _shapes.Count;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _shapes.Parent;

    /// <summary>
    /// 根据索引获取形状
    /// </summary>
    public IPowerPointShape this[int index]
    {
        get
        {
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

                try
                {
                    var shape = _shapes[index];
                    return new PowerPointShape(shape);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to get shape at index {index}.", ex);
                }
            }
        }
    }

    /// <summary>
    /// 根据名称获取形状
    /// </summary>
    public IPowerPointShape this[string name]
    {
        get
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Shape name cannot be null or empty.", nameof(name));

            try
            {
                var shape = _shapes[name];
                return new PowerPointShape(shape);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get shape with name '{name}'.", ex);
            }
        }
    }


    /// <summary>
    /// 获取主标题占位符
    /// </summary>
    public IPowerPointShape Title
    {
        get
        {
            try
            {
                var titleShape = _shapes.Title;
                return titleShape != null ? new PowerPointShape(titleShape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="shapes">COM Shapes 对象</param>
    internal PowerPointShapes(MsPowerPoint.Shapes shapes)
    {
        _shapes = shapes ?? throw new ArgumentNullException(nameof(shapes));
        _disposedValue = false;
    }


    /// <summary>
    /// 添加形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的形状</returns>
    public IPowerPointShape AddShape(MsoAutoShapeType type, double left, double top, double width, double height)
    {
        try
        {
            var shape = _shapes.AddShape((MsCore.MsoAutoShapeType)type, (float)left, (float)top, (float)width, (float)height);
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add shape.", ex);
        }
    }

    /// <summary>
    /// 添加文本框
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的文本框</returns>
    public IPowerPointShape AddTextbox(MsoTextOrientation orientation, double left, double top, double width, double height)
    {
        try
        {
            var shape = _shapes.AddTextbox((MsCore.MsoTextOrientation)orientation, (float)left, (float)top, (float)width, (float)height);
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add textbox.", ex);
        }
    }

    /// <summary>
    /// 添加图片
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的图片形状</returns>
    public IPowerPointShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument, double left, double top, double width, double height)
    {
        if (!System.IO.File.Exists(fileName))
            throw new System.IO.FileNotFoundException("Picture file not found.", fileName);

        try
        {
            var shape = _shapes.AddPicture(fileName,
                linkToFile ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                saveWithDocument ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                (float)left, (float)top, (float)width, (float)height);
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add picture from '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 添加图表
    /// </summary>
    /// <param name="type">图表类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的图表形状</returns>
    public IPowerPointShape AddChart(MsoChartType type, double left, double top, double width, double height)
    {
        try
        {
            var shape = _shapes.AddChart((MsCore.XlChartType)type, (float)left, (float)top, (float)width, (float)height);
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add chart.", ex);
        }
    }

    /// <summary>
    /// 添加表格
    /// </summary>
    /// <param name="numRows">行数</param>
    /// <param name="numColumns">列数</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的表格形状</returns>
    public IPowerPointShape AddTable(int numRows, int numColumns, double left, double top, double width, double height)
    {
        if (numRows <= 0 || numColumns <= 0)
            throw new ArgumentException("Rows and columns must be greater than zero.");

        try
        {
            var shape = _shapes.AddTable(numRows, numColumns, (float)left, (float)top, (float)width, (float)height);
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add table.", ex);
        }
    }

    /// <summary>
    /// 添加智能图形
    /// </summary>
    /// <param name="smartArtType">智能图形类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的智能图形形状</returns>
    public IPowerPointShape AddSmartArt(object smartArtType, double left, double top, double width, double height)
    {
        try
        {
            var shape = _shapes.AddSmartArt((MsCore.SmartArtLayout)smartArtType, (float)left, (float)top, (float)width, (float)height);
            return shape != null ? new PowerPointShape(shape) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add SmartArt.", ex);
        }
    }

    public IPowerPointShape AddOLEObject(
        float Left = 0f, float Top = 0f,
        float Width = -1f, float Height = -1f,
        string ClassName = "", string FileName = "", bool DisplayAsIcon = false,
        string IconFileName = "", int IconIndex = 0,
        string IconLabel = "", bool Link = false)
    {
        var shape = _shapes.AddOLEObject(Left, Top, Width, Height,
            ClassName, FileName, DisplayAsIcon ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
          IconFileName, IconIndex, IconLabel,
          Link ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
        return shape != null ? new PowerPointShape(shape) : null;
    }

    /// <summary>
    /// 获取形状范围
    /// </summary>
    /// <param name="index">索引或名称数组</param>
    /// <returns>形状范围对象</returns>
    public IPowerPointShapeRange Range(object index)
    {
        try
        {
            var shapeRange = _shapes.Range(index);
            return new PowerPointShapeRange(shapeRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get shape range.", ex);
        }
    }

    /// <summary>
    /// 根据条件查找形状
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的形状列表</returns>
    public IEnumerable<IPowerPointShape> Find(Func<IPowerPointShape, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        var results = new List<IPowerPointShape>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var shape = this[i];
                    if (predicate(shape))
                    {
                        results.Add(shape);
                    }
                }
                catch
                {
                    // 忽略获取或判断单个形状失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find shapes.", ex);
        }
        return results;
    }

    /// <summary>
    /// 按类型查找形状
    /// </summary>
    /// <param name="shapeType">形状类型</param>
    /// <returns>指定类型的形状列表</returns>
    public IEnumerable<IPowerPointShape> FindByType(MsoShapeType shapeType)
    {
        return Find(shape => shape.Type == shapeType);
    }

    /// <summary>
    /// 按名称查找形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的形状列表</returns>
    public IEnumerable<IPowerPointShape> FindByName(string name, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(name))
            return new List<IPowerPointShape>();

        return Find(shape =>
        {
            var shapeName = shape.Name ?? string.Empty;
            return matchCase ? shapeName == name : shapeName.Equals(name, StringComparison.OrdinalIgnoreCase);
        });
    }

    /// <summary>
    /// 删除所有形状
    /// </summary>
    public void Delete()
    {
        try
        {
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    _shapes[i].Delete();
                }
                catch
                {
                    // 忽略删除单个形状失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete all shapes.", ex);
        }
    }

    /// <summary>
    /// 删除指定索引的形状
    /// </summary>
    /// <param name="index">形状索引</param>
    public void Delete(int index)
    {
        try
        {
            this[index].Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete shape at index {index}.", ex);
        }
    }

    /// <summary>
    /// 删除指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    public void Delete(string name)
    {
        try
        {
            this[name].Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete shape with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 选择所有形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void SelectAll(bool replace = true)
    {
        try
        {
            _shapes.SelectAll();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select all shapes.", ex);
        }
    }

    /// <summary>
    /// 取消选择所有形状
    /// </summary>
    public void DeselectAll()
    {
        try
        {
            // PowerPoint 中没有直接的取消选择方法，这里通过选择幻灯片来实现
            var slide = Parent as MsPowerPoint.Slide;
            slide?.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to deselect all shapes.", ex);
        }
    }

    /// <summary>
    /// 对齐所有形状
    /// </summary>
    /// <param name="alignCmd">对齐命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片对齐</param>
    public void AlignAll(int alignCmd, bool relativeToSlide = false)
    {
        try
        {
            if (Count > 0)
            {
                var range = _shapes.Range();
                range.Align((MsCore.MsoAlignCmd)alignCmd,
                    relativeToSlide ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to align all shapes.", ex);
        }
    }

    /// <summary>
    /// 分布所有形状
    /// </summary>
    /// <param name="distributeCmd">分布命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片分布</param>
    public void DistributeAll(int distributeCmd, bool relativeToSlide = false)
    {
        try
        {
            if (Count > 0)
            {
                var range = _shapes.Range();
                range.Distribute((MsCore.MsoDistributeCmd)distributeCmd,
                    relativeToSlide ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to distribute all shapes.", ex);
        }
    }

    /// <summary>
    /// 组合所有形状
    /// </summary>
    /// <returns>组合后的形状</returns>
    public IPowerPointShape GroupAll()
    {
        try
        {
            if (Count > 0)
            {
                var range = _shapes.Range();
                var groupedShape = range.Group();
                return new PowerPointShape(groupedShape);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to group all shapes.", ex);
        }
    }

    /// <summary>
    /// 获取占位符
    /// </summary>
    /// <param name="index">占位符索引</param>
    /// <returns>占位符形状</returns>
    public IPowerPointShape Placeholders(int index)
    {
        try
        {
            var placeholder = _shapes.Placeholders[index];
            return new PowerPointShape(placeholder);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get placeholder at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取所有占位符
    /// </summary>
    public IEnumerable<IPowerPointShape> GetAllPlaceholders()
    {
        var placeholders = new List<IPowerPointShape>();
        try
        {
            var placeholdersCollection = _shapes.Placeholders;
            for (int i = 1; i <= placeholdersCollection.Count; i++)
            {
                try
                {
                    var placeholder = placeholdersCollection[i];
                    placeholders.Add(new PowerPointShape(placeholder));
                }
                catch
                {
                    // 忽略获取单个占位符失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all placeholders.", ex);
        }
        return placeholders;
    }

    /// <summary>
    /// 按Z轴顺序排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的形状列表</returns>
    public IEnumerable<IPowerPointShape> OrderByZOrder(bool ascending = true)
    {
        var shapes = new List<IPowerPointShape>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    shapes.Add(this[i]);
                }
                catch
                {
                    // 忽略获取单个形状失败的情况
                    continue;
                }
            }

            return ascending
                ? shapes.OrderBy(s => s.ZOrderPosition)
                : shapes.OrderByDescending(s => s.ZOrderPosition);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to order shapes by Z-order.", ex);
        }
    }

    /// <summary>
    /// 按名称排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的形状列表</returns>
    public IEnumerable<IPowerPointShape> OrderByName(bool ascending = true)
    {
        var shapes = new List<IPowerPointShape>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    shapes.Add(this[i]);
                }
                catch
                {
                    // 忽略获取单个形状失败的情况
                    continue;
                }
            }

            return ascending
                ? shapes.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
                : shapes.OrderByDescending(s => s.Name, StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to order shapes by name.", ex);
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

    public IEnumerator<IPowerPointShape> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

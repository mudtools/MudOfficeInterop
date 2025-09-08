//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shapes 的实现类。
/// </summary>
internal class WordShapes : IWordShapes
{
    private MsWord.Shapes _shapes;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="shapes">原始 COM Shapes 对象。</param>
    internal WordShapes(MsWord.Shapes shapes)
    {
        _shapes = shapes ?? throw new ArgumentNullException(nameof(shapes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _shapes != null ? new WordApplication(_shapes.Application) : null;

    /// <inheritdoc/>
    public object Parent => _shapes?.Parent;

    /// <inheritdoc/>
    public int Count => _shapes?.Count ?? 0;

    /// <inheritdoc/>
    public IWordShape this[int index]
    {
        get
        {
            if (_shapes == null || index < 1 || index > Count)
                return null;

            var shape = _shapes[index];
            return new WordShape(shape);
        }
    }

    /// <inheritdoc/>
    public IWordShape this[string name]
    {
        get
        {
            if (_shapes == null || string.IsNullOrWhiteSpace(name))
                return null;

            try
            {
                var shape = _shapes[name];
                return shape != null ? new WordShape(shape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordShape? AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height)
    {
        if (_shapes == null)
        {
            throw new ObjectDisposedException(nameof(WordShapes));
        }

        try
        {
            var shape = _shapes.AddTextbox((MsCore.MsoTextOrientation)(int)orientation, left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加文本框形状。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordShape? AddRectangle(float left, float top, float width, float height)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(WordShapes));

        try
        {
            var shape = _shapes.AddShape((int)MsCore.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加矩形形状。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordShape? AddLine(float beginX, float beginY, float endX, float endY)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(WordShapes));

        try
        {
            var shape = _shapes.AddLine(beginX, beginY, endX, endY);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加线条形状。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordShape? AddPicture(string fileName, bool linkToFile, bool saveWithDocument, float left, float top, float width, float height)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(WordShapes));

        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("文件名不能为空。", nameof(fileName));

        try
        {
            var shape = _shapes.AddPicture(fileName, linkToFile ? 1 : 0, saveWithDocument ? 1 : 0, left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加图片形状 '{fileName}'。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordShape? AddChart(MsoChartType type, float left, float top,
        float width, float height)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(WordShapes));

        try
        {
            var shape = _shapes.AddChart((MsCore.XlChartType)(int)type, left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加图表形状。", ex);
        }
    }

    public IWordShape? AddOLEObject(string? classType = null, string? fileName = null, bool? linkToFile = false,
        bool? displayAsIcon = false, string? iconFileName = null, int? iconIndex = null, string? iconLabel = null,
        float? left = null, float? top = null, float? width = null, float? height = null, object? anchor = null)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(WordShapes));

        var shape = _shapes.AddOLEObject(
            classType.ComArgsVal(), fileName.ComArgsVal(), linkToFile.ComArgsVal(),
            displayAsIcon.ComArgsVal(), iconFileName.ComArgsVal(), iconIndex.ComArgsVal(),
            iconLabel.ComArgsVal(), left.ComArgsVal(), top.ComArgsVal(),
            width.ComArgsVal(), height.ComArgsVal(), anchor != null ? anchor : Type.Missing);

        return shape != null ? new WordShape(shape) : null;
    }


    /// <inheritdoc/>
    public bool Contains(string name)
    {
        if (_shapes == null || string.IsNullOrWhiteSpace(name))
            return false;

        return _shapes[name] != null;
    }

    /// <inheritdoc/>
    public List<string> GetAllShapeNames()
    {
        var names = new List<string>();

        if (_shapes == null)
            return names;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes[i];
            if (shape?.Name != null)
            {
                names.Add(shape.Name);
            }
        }

        return names;
    }

    /// <inheritdoc/>
    public List<string> GetShapeNamesByType(MsoShapeType shapeType)
    {
        var names = new List<string>();

        if (_shapes == null)
            return names;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes[i];
            if (shape != null && shape.Type == (MsCore.MsoShapeType)(int)shapeType)
            {
                names.Add(shape.Name);
            }
        }

        return names;
    }

    /// <inheritdoc/>
    public bool DeleteShape(string name)
    {
        if (_shapes == null || string.IsNullOrWhiteSpace(name))
            return false;

        var shape = _shapes[name];
        if (shape != null)
        {
            shape.Delete();
            return true;
        }
        return false;
    }

    /// <inheritdoc/>
    public void DeleteAll()
    {
        if (_shapes == null)
            return;

        // 从后往前删除，避免索引变化问题
        for (int i = Count; i >= 1; i--)
        {
            _shapes[i]?.Delete();
        }
    }

    /// <inheritdoc/>
    public void SelectAll()
    {
        if (_shapes == null)
            return;
        _shapes.SelectAll();
    }

    /// <inheritdoc/>
    public IWordShapeRange GetShapesInRange(IWordRange range)
    {
        if (_shapes == null || range == null)
            return null;

        try
        {
            var wordRange = (range as WordRange)?._range;
            if (wordRange != null)
            {
                var rangeShapes = wordRange.ShapeRange;
                return new WordShapeRange(rangeShapes);
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public int GetCountByType(MsoShapeType shapeType)
    {
        if (_shapes == null)
            return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes[i];
            if (shape != null && shape.Type == (MsCore.MsoShapeType)(int)shapeType)
            {
                count++;
            }
        }
        return count;
    }

    #endregion

    #region IEnumerable<IWordShape> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordShape> GetEnumerator()
    {
        if (_shapes == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes[i];
            if (shape != null)
                yield return new WordShape(shape);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shapes != null)
        {
            Marshal.ReleaseComObject(_shapes);
            _shapes = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
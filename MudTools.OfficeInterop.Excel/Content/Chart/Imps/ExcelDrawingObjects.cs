//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelDrawingObjects : IExcelDrawingObjects
{
    private MsExcel.DrawingObjects _drawingObjects;
    private bool _disposedValue;

    public int Count => _drawingObjects.Count;

    public int ZOrder => _drawingObjects.ZOrder;

    public double Width
    {
        get => _drawingObjects.Width;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Width = value;
        }
    }
    public double Height
    {
        get => _drawingObjects.Height;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Height = value;
        }
    }

    public double Top
    {
        get => _drawingObjects.Top;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Top = value;
        }
    }
    public double Left
    {
        get => _drawingObjects.Left;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Left = value;
        }
    }
    public bool Locked
    {
        get => _drawingObjects.Locked;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Locked = value;
        }
    }


    public bool Enabled
    {
        get => _drawingObjects.Enabled;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Enabled = value;
        }
    }


    public bool Visible
    {
        get => _drawingObjects.Visible;
        set
        {
            if (_drawingObjects != null)
                _drawingObjects.Visible = value;
        }
    }



    public IExcelShapeRange ShapeRange =>
        _drawingObjects != null ? new ExcelShapeRange(_drawingObjects.ShapeRange) : null;

    public IExcelDrawing this[int index] => new ExcelDrawing(_drawingObjects.Item(index) as MsExcel.Drawing);

    public IExcelDrawing this[string name] => new ExcelDrawing(_drawingObjects.Item(name) as MsExcel.Drawing);

    internal ExcelDrawingObjects(MsExcel.DrawingObjects drawingObjects)
    {
        _drawingObjects = drawingObjects ?? throw new ArgumentNullException(nameof(drawingObjects));
        _disposedValue = false;
    }

    public IExcelDrawing GetItem(object index)
    {
        try
        {
            var drawing = _drawingObjects.Item(index) as MsExcel.Drawing;
            return drawing != null ? new ExcelDrawing(drawing) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }


    public IExcelDrawing FindByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("对象名称不能为空。", nameof(name));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var drawing = this[i];
                if (string.Equals(drawing.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return drawing;
                }
            }
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public void Remove(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("对象名称不能为空。", nameof(name));

        try
        {
            var drawing = FindByName(name);
            if (drawing != null)
            {
                drawing.Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除绘图对象: {name}", ex);
        }
    }

    public void Clear()
    {
        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    this[i].Delete();
                }
                catch (COMException)
                {
                    // 继续删除其他对象
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除所有绘图对象。", ex);
        }
    }

    public void SelectAll()
    {
        try
        {
            _drawingObjects.Select();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择所有绘图对象。", ex);
        }
    }

    public IEnumerable<IExcelDrawing> VisibleItems
    {
        get
        {
            var result = new List<IExcelDrawing>();
            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    var drawing = this[i];
                    if (drawing.Visible)
                    {
                        result.Add(drawing);
                    }
                }
            }
            catch (COMException)
            {
                // 忽略异常，返回已找到的结果
            }
            return result;
        }
    }

    public IEnumerable<IExcelDrawing> LockedItems
    {
        get
        {
            var result = new List<IExcelDrawing>();
            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    var drawing = this[i];
                    if (drawing.Locked)
                    {
                        result.Add(drawing);
                    }
                }
            }
            catch (COMException)
            {

            }
            return result;
        }
    }

    public IEnumerator<IExcelDrawing> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _drawingObjects != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_drawingObjects) > 0) { }
            }
            catch { }
            _drawingObjects = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

}
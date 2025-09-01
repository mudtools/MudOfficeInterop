//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// 对 Microsoft.Office.Core.Shapes 的二次封装实现类。
/// 提供对 Shapes 集合的安全访问和管理。
/// </summary>
internal class OfficeShapes : IOfficeShapes
{
    private MsCore.Shapes _shapes;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 OfficeShapes 类的新实例
    /// </summary>
    /// <param name="shapes">原始的 COM 形状集合对象</param>
    internal OfficeShapes(MsCore.Shapes shapes)
    {
        _shapes = shapes ?? throw new ArgumentNullException(nameof(shapes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _shapes?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeShape this[int index]
    {
        get
        {
            if (_shapes == null || index < 1 || index > Count)
                return null;

            try
            {
                var shape = _shapes.Item(index);
                return shape != null ? new OfficeShape(shape) : null;
            }
            catch (ArgumentException)
            {
                // 索引超出范围时返回null
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IOfficeShape this[string name]
    {
        get
        {
            if (_shapes == null || string.IsNullOrWhiteSpace(name))
                return null;

            try
            {
                var shape = _shapes.Item(name);
                return shape != null ? new OfficeShape(shape) : null;
            }
            catch (ArgumentException)
            {
                // 名称不存在时返回null
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficeShape AddShape(MsoAutoShapeType type, float left, float top, float width, float height)
    {
        if (_shapes == null)
            return null;

        var shape = _shapes.AddShape((MsCore.MsoAutoShapeType)(int)type, left, top, width, height);
        return shape != null ? new OfficeShape(shape) : null;
    }

    /// <inheritdoc/>
    public IOfficeShape AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height)
    {
        if (_shapes == null)
            return null;

        var textbox = _shapes.AddTextbox((MsCore.MsoTextOrientation)(int)orientation, left, top, width, height);
        return textbox != null ? new OfficeShape(textbox) : null;
    }

    /// <inheritdoc/>
    public void DeleteAll()
    {
        if (_shapes == null)
            return;

        // 从后向前删除所有形状，避免索引变化导致的问题
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                _shapes.Item(i).Delete();
            }
            catch
            {
                // 忽略删除过程中可能出现的异常
            }
        }
    }

    /// <inheritdoc/>
    public IOfficeShape SelectByName(string name)
    {
        if (_shapes == null || string.IsNullOrWhiteSpace(name))
            return null;

        try
        {
            _shapes.Item(name).Select();
            return this[name];
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public IEnumerable<IOfficeShape> GetRange()
    {
        if (_shapes == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes.Item(i);
            if (shape != null)
                yield return new OfficeShape(shape);
        }
    }

    #endregion

    #region IEnumerable<IOfficeShape> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeShape> GetEnumerator()
    {
        if (_shapes == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapes.Item(i);
            if (shape != null)
                yield return new OfficeShape(shape);
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
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在处置</param>
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
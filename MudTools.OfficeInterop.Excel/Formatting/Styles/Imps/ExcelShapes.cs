//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Shapes 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Shapes 对象的安全访问和资源管理
/// </summary>
internal class ExcelShapes : IExcelShapes
{
    /// <summary>
    /// 底层的 COM Shapes 集合对象
    /// </summary>
    private MsExcel.Shapes _shapes;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelShapes 实例
    /// </summary>
    /// <param name="shapes">底层的 COM Shapes 集合对象</param>
    internal ExcelShapes(MsExcel.Shapes shapes)
    {
        _shapes = shapes;
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
                // 释放所有子形状对象
                for (int i = 1; i <= Count; i++)
                {
                    var shape = this[i] as ExcelShape;
                    shape?.Dispose();
                }

                // 释放底层COM对象
                if (_shapes != null)
                    Marshal.ReleaseComObject(_shapes);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _shapes = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取形状集合中的形状数量
    /// </summary>
    public int Count => _shapes?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的形状对象
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    public IExcelShape this[int index]
    {
        get
        {
            if (_shapes == null || index < 1 || index > Count)
                return null;

            var shape = _shapes.Item(index);
            return shape != null ? new ExcelShape(shape) : null;
        }
    }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    public IExcelShape this[string name]
    {
        get
        {
            if (_shapes == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var shape = _shapes.Item(name);
                return shape != null ? new ExcelShape(shape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 添加文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape AddTextbox(int orientation, double left, double top, double width, double height)
    {
        if (_shapes == null)
            return null;

        return _shapes.AddTextbox((MsCore.MsoTextOrientation)orientation, (float)left, (float)top, (float)width, (float)height) is MsExcel.Shape shape ? new ExcelShape(shape) : null;
    }

    /// <summary>
    /// 添加矩形形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape AddRectangle(double left, double top, double width, double height)
    {
        if (_shapes == null)
            return null;

        var shape = _shapes.AddShape(MsCore.MsoAutoShapeType.msoShapeRectangle, (float)left, (float)top, (float)width, (float)height) as MsExcel.Shape;
        return shape != null ? new ExcelShape(shape) : null;
    }

    /// <summary>
    /// 添加椭圆形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape AddEllipse(double left, double top, double width, double height)
    {
        if (_shapes == null)
            return null;

        var shape = _shapes.AddShape(MsCore.MsoAutoShapeType.msoShapeOval, (float)left, (float)top, (float)width, (float)height) as MsExcel.Shape;
        return shape != null ? new ExcelShape(shape) : null;
    }

    /// <summary>
    /// 添加线条形状
    /// </summary>
    /// <param name="x1">起点X坐标</param>
    /// <param name="y1">起点Y坐标</param>
    /// <param name="x2">终点X坐标</param>
    /// <param name="y2">终点Y坐标</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape AddLine(double x1, double y1, double x2, double y2)
    {
        if (_shapes == null)
            return null;

        var shape = _shapes.AddLine((float)x1, (float)y1, (float)x2, (float)y2) as MsExcel.Shape;
        return shape != null ? new ExcelShape(shape) : null;
    }

    /// <summary>
    /// 添加图片形状
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape? AddPicture(
        string filename,
        bool linkToFile,
        bool saveWithDocument,
        double left,
        double top,
        double width,
        double height)
    {
        if (_shapes == null || string.IsNullOrEmpty(filename))
            return null;

        var shape = _shapes.AddPicture(filename,
            linkToFile ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            saveWithDocument ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            (float)left, (float)top, (float)width, (float)height);
        return shape != null ? new ExcelShape(shape) : null;
    }

    public IExcelShapeRange? Range(string index)
    {
        return _shapes != null ? new ExcelShapeRange(_shapes.Range[index]) : null;
    }

    /// <summary>
    /// 选择所有形状
    /// </summary>
    public void SelectAll()
    {
        _shapes?.SelectAll();
    }

    /// <summary>
    /// 删除所有形状
    /// </summary>
    public void DeleteAll()
    {
        if (_shapes == null) return;

        // 从后往前删除，避免索引变化问题
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                _shapes.Item(i).Delete();
            }
            catch
            {
                // 忽略删除过程中的异常
            }
        }
    }

    public IEnumerator<IExcelShape> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return new ExcelShape(_shapes.Item(i));
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
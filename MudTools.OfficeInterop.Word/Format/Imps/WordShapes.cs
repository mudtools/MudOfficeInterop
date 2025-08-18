//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Word.Shapes 的封装实现类
/// </summary>
internal class WordShapes : IWordShapes
{
    #region 属性封装

    /// <summary>
    /// 获取形状的数量
    /// </summary>
    public int Count => _shapes.Count;

    /// <summary>
    /// 获取指定索引的形状对象
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    public IWordShape this[int index] => new WordShape(_shapes[index]);

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    public IWordShape this[string name] => new WordShape(_shapes[name]);

    #endregion

    #region 构造函数与私有字段

    private MsWord.Shapes _shapes;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 WordShapes 实例
    /// </summary>
    /// <param name="shapes">原始 COM Shapes 对象</param>
    internal WordShapes(MsWord.Shapes shapes)
    {
        _shapes = shapes ?? throw new ArgumentNullException(nameof(shapes));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    public IWordShape AddOLEObject(ref object ClassType,
        ref object FileName,
        ref object LinkToFile,
        ref object DisplayAsIcon,
        ref object IconFileName,
        ref object IconIndex,
        ref object IconLabel,
        ref object Left,
        ref object Top,
        ref object Width,
        ref object Height,
        ref object Anchor)
    {
        var shape = _shapes.AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height, Anchor);
        return new WordShape(shape);
    }


    /// <summary>
    /// 添加文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">上边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IWordShape AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height)
    {
        try
        {
            var shape = _shapes.AddTextbox((MsCore.MsoTextOrientation)(int)orientation,
                left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加文本框形状", ex);
        }
    }

    /// <summary>
    /// 添加图片形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">上边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IWordShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument,
                                double left, double top, double width, double height)
    {
        try
        {
            var shape = _shapes.AddPicture(fileName, linkToFile, saveWithDocument,
                                         left, top, width, height);
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加图片形状", ex);
        }
    }

    /// <summary>
    /// 根据索引删除形状
    /// </summary>
    /// <param name="index">要删除的形状索引</param>
    public void Delete(int index)
    {
        try
        {
            _shapes[index].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除索引为 {index} 的形状", ex);
        }
    }

    /// <summary>
    /// 根据名称删除形状
    /// </summary>
    /// <param name="name">要删除的形状名称</param>
    public void Delete(string name)
    {
        try
        {
            _shapes[name].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除名称为 {name} 的形状", ex);
        }
    }

    /// <summary>
    /// 删除所有形状
    /// </summary>
    public void DeleteAll()
    {
        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                _shapes[i].Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除所有形状", ex);
        }
    }



    /// <summary>
    /// 查找指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象，如果未找到则返回null</returns>
    public IWordShape FindByName(string name)
    {
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var shape = _shapes[i];
                if (string.Equals(shape.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return new WordShape(shape);
                }
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region IDisposable 模式实现

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否显式调用 Dispose()</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shapes != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_shapes) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _shapes = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 显式释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// 获取所有形状的枚举器
    /// </summary>
    public IEnumerator<IWordShape> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
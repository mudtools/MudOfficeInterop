//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Word.InlineShapes 的封装实现类
/// </summary>
internal class WordInlineShapes : IWordInlineShapes
{
    #region 属性封装

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _inlineShapes != null ? new WordApplication(_inlineShapes.Application) : null;

    /// <summary>
    /// 获取内嵌形状的数量
    /// </summary>
    public int Count => Convert.ToInt32(_inlineShapes.Count);

    /// <summary>
    /// 获取指定索引的内嵌形状对象
    /// </summary>
    /// <param name="index">内嵌形状索引（从1开始）</param>
    /// <returns>内嵌形状对象</returns>
    public IWordInlineShape this[int index] => new WordInlineShape(_inlineShapes[index]);

    /// <summary>
    /// 获取集合的父对象（伪代码）
    /// </summary>
    public object Parent => _inlineShapes.Parent;

    #endregion

    #region 构造函数与私有字段

    private MsWord.InlineShapes _inlineShapes;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 WordInlineShapes 实例
    /// </summary>
    /// <param name="inlineShapes">原始 COM InlineShapes 对象</param>
    internal WordInlineShapes(MsWord.InlineShapes inlineShapes)
    {
        _inlineShapes = inlineShapes ?? throw new ArgumentNullException(nameof(inlineShapes));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    /// <summary>
    /// 添加图片内嵌形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <returns>新创建的内嵌形状对象</returns>
    public IWordInlineShape AddPicture(string fileName,
        bool linkToFile = false, bool saveWithDocument = true)
    {
        try
        {
            object linkToFileObj = linkToFile;
            object saveWithDocumentObj = saveWithDocument;
            var inlineShape = _inlineShapes.AddPicture(fileName,
                ref linkToFileObj,
                ref saveWithDocumentObj);
            return new WordInlineShape(inlineShape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加图片内嵌形状", ex);
        }
    }

    /// <summary>
    /// 添加OLE对象内嵌形状
    /// </summary>
    /// <param name="classType">OLE对象类类型</param>
    /// <param name="fileName">文件名</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="displayAsIcon">是否以图标显示</param>
    /// <param name="iconFileName">图标文件路径</param>
    /// <param name="iconIndex">图标索引</param>
    /// <param name="iconLabel">图标标签</param>
    /// <returns>新创建的内嵌形状对象</returns>
    public IWordInlineShape AddOLEObject(string classType = null, string fileName = null, bool linkToFile = false,
                                        bool displayAsIcon = false, string iconFileName = null, int iconIndex = 0,
                                        string iconLabel = null)
    {
        try
        {
            var inlineShape = _inlineShapes.AddOLEObject(classType, fileName, linkToFile, displayAsIcon,
                                                       iconFileName, iconIndex, iconLabel);
            return new WordInlineShape(inlineShape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加OLE对象内嵌形状", ex);
        }
    }

    /// <summary>
    /// 添加水平线内嵌形状
    /// </summary>
    /// <param name="fileName">水平线文件路径</param>
    /// <param name="Range">是否链接到文件</param>
    /// <returns>新创建的内嵌形状对象</returns>
    public IWordInlineShape AddHorizontalLine(string fileName = null, object Range = null)
    {
        try
        {
            var inlineShape = _inlineShapes.AddHorizontalLine(fileName, Range);
            return new WordInlineShape(inlineShape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加水平线内嵌形状", ex);
        }
    }


    /// <summary>
    /// 添加图表内嵌形状
    /// </summary>
    /// <param name="style">图表样式</param>
    /// <returns>新创建的内嵌形状对象</returns>
    public IWordInlineShape AddChart(MsoChartType style = MsoChartType.xlArea)
    {
        try
        {
            var inlineShape = _inlineShapes.AddChart((MsCore.XlChartType)(int)style);
            return new WordInlineShape(inlineShape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加图表内嵌形状", ex);
        }
    }

    /// <summary>
    /// 根据索引删除内嵌形状
    /// </summary>
    /// <param name="index">要删除的内嵌形状索引</param>
    public void Delete(int index)
    {
        try
        {
            _inlineShapes[index].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除索引为 {index} 的内嵌形状", ex);
        }
    }

    /// <summary>
    /// 删除所有内嵌形状
    /// </summary>
    public void DeleteAll()
    {
        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                _inlineShapes[i].Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除所有内嵌形状", ex);
        }
    }



    /// <summary>
    /// 查找指定类型的内嵌形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <returns>符合条件的内嵌形状集合</returns>
    public IEnumerable<IWordInlineShape> FindByType(WdInlineShapeType type)
    {
        for (int i = 1; i <= Count; i++)
        {
            var shape = this[i];
            if (shape.Type == type)
            {
                yield return shape;
            }
            else
            {
                shape.Dispose();
            }
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

        if (disposing && _inlineShapes != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_inlineShapes) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _inlineShapes = null;
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


    public IEnumerator<IWordInlineShape> GetEnumerator()
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

    #endregion
}
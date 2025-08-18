//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Word.InlineShape 的封装实现类
/// </summary>
internal class WordInlineShape : IWordInlineShape
{
    #region 属性封装

    /// <summary>
    /// 获取内嵌形状的类型
    /// </summary>
    public int Type => Convert.ToInt32(_inlineShape.Type);

    /// <summary>
    /// 获取或设置内嵌形状的替代文本
    /// </summary>
    public string AlternativeText
    {
        get => _inlineShape.AlternativeText;
        set => _inlineShape.AlternativeText = value;
    }

    /// <summary>
    /// 获取或设置内嵌形状的高度
    /// </summary>
    public float Height
    {
        get => _inlineShape.Height;
        set => _inlineShape.Height = value;
    }

    /// <summary>
    /// 获取或设置内嵌形状的宽度
    /// </summary>
    public float Width
    {
        get => _inlineShape.Width;
        set => _inlineShape.Width = value;
    }

    /// <summary>
    /// 获取内嵌形状的范围（伪代码）
    /// </summary>
    public object Range => _inlineShape.Range;

    /// <summary>
    /// 获取内嵌形状的父对象（伪代码）
    /// </summary>
    public object Parent => _inlineShape.Parent;

    /// <summary>
    /// 获取内嵌形状的OLE格式（伪代码）
    /// </summary>
    public object OLEFormat => _inlineShape.OLEFormat;

    /// <summary>
    /// 获取内嵌形状的链接格式（伪代码）
    /// </summary>
    public object LinkFormat => _inlineShape.LinkFormat;


    /// <summary>
    /// 获取内嵌形状的填充属性（伪代码）
    /// </summary>
    public object Fill => _inlineShape.Fill;

    /// <summary>
    /// 获取内嵌形状的线条属性（伪代码）
    /// </summary>
    public object Line => _inlineShape.Line;


    /// <summary>
    /// 锁定内嵌形状的比例
    /// </summary>
    public bool LockAspectRatio
    {
        get => _inlineShape.LockAspectRatio == MsCore.MsoTriState.msoTrue ? true : false;
        set => _inlineShape.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取内嵌形状是否为图片类型
    /// </summary>
    public bool IsPicture => Type == 3; // wdInlineShapePicture = 3

    /// <summary>
    /// 获取内嵌形状是否为OLE对象类型
    /// </summary>
    public bool IsOLEObject => Type == 1; // wdInlineShapeEmbeddedOLEObject = 1

    #endregion

    #region 构造函数与私有字段

    private MsWord.InlineShape _inlineShape;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 WordInlineShape 实例
    /// </summary>
    /// <param name="inlineShape">原始 COM InlineShape 对象</param>
    internal WordInlineShape(MsWord.InlineShape inlineShape)
    {
        _inlineShape = inlineShape ?? throw new ArgumentNullException(nameof(inlineShape));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    /// <summary>
    /// 删除当前内嵌形状
    /// </summary>
    public void Delete()
    {
        try
        {
            _inlineShape.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除内嵌形状", ex);
        }
    }

    /// <summary>
    /// 复制当前内嵌形状
    /// </summary>
    public void Copy()
    {
        try
        {
            _inlineShape.Range.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制内嵌形状", ex);
        }
    }

    /// <summary>
    /// 剪切当前内嵌形状
    /// </summary>
    public void Cut()
    {
        try
        {
            _inlineShape.Range.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切内嵌形状", ex);
        }
    }

    /// <summary>
    /// 选择当前内嵌形状
    /// </summary>
    public void Select()
    {
        try
        {
            _inlineShape.Range.Select();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择内嵌形状", ex);
        }
    }

    /// <summary>
    /// 将内嵌形状转换为浮动形状
    /// </summary>
    /// <returns>转换后的浮动形状对象</returns>
    public IWordShape ConvertToShape()
    {
        try
        {
            var shape = _inlineShape.ConvertToShape();
            return new WordShape(shape);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法转换为浮动形状", ex);
        }
    }

    /// <summary>
    /// 更新链接的内嵌形状
    /// </summary>
    public void Update()
    {
        try
        {
            _inlineShape.LinkFormat?.Update();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法更新链接的内嵌形状", ex);
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

        if (disposing && _inlineShape != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_inlineShape) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _inlineShape = null;
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

    #endregion
}
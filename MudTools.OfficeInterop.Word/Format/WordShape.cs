//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Word.Shape 的封装实现类
/// </summary>
internal class WordShape : IWordShape
{
    #region 属性封装

    /// <summary>
    /// 获取或设置形状的名称
    /// </summary>
    public string Name
    {
        get => _shape.Name?.ToString();
        set => _shape.Name = value;
    }

    /// <summary>
    /// 获取或设置形状的左边距
    /// </summary>
    public float Left
    {
        get => _shape.Left;
        set => _shape.Left = value;
    }

    /// <summary>
    /// 获取或设置形状的上边距
    /// </summary>
    public float Top
    {
        get => _shape.Top;
        set => _shape.Top = value;
    }

    /// <summary>
    /// 获取或设置形状的宽度
    /// </summary>
    public float Width
    {
        get => _shape.Width;
        set => _shape.Width = value;
    }

    /// <summary>
    /// 获取或设置形状的高度
    /// </summary>
    public float Height
    {
        get => _shape.Height;
        set => _shape.Height = value;
    }

    /// <summary>
    /// 获取或设置形状是否可见
    /// </summary>
    public bool Visible
    {
        get => _shape.Visible == MsCore.MsoTriState.msoTrue ? true : false;
        set => _shape.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }


    /// <summary>
    /// 获取或设置形状的替代文本
    /// </summary>
    public string AlternativeText
    {
        get => _shape.AlternativeText;
        set => _shape.AlternativeText = value;
    }

    /// <summary>
    /// 获取形状的文本框架（伪代码）
    /// </summary>
    public object TextFrame => _shape.TextFrame;

    /// <summary>
    /// 获取形状的填充属性（伪代码）
    /// </summary>
    public object Fill => _shape.Fill;

    /// <summary>
    /// 获取形状的线条属性（伪代码）
    /// </summary>
    public object Line => _shape.Line;


    /// <summary>
    /// 获取形状的Z轴顺序位置
    /// </summary>
    public int ZOrderPosition => _shape.ZOrderPosition;


    public IWordOLEFormat OLEFormat => new WordOLEFormat(_shape.OLEFormat);

    #endregion

    #region 构造函数与私有字段

    private MsWord.Shape _shape;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 WordShape 实例
    /// </summary>
    /// <param name="shape">原始 COM Shape 对象</param>
    internal WordShape(MsWord.Shape shape)
    {
        _shape = shape ?? throw new ArgumentNullException(nameof(shape));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    /// <summary>
    /// 删除当前形状
    /// </summary>
    public void Delete()
    {
        try
        {
            _shape.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除形状", ex);
        }
    }


    /// <summary>
    /// 选择当前形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        try
        {
            _shape.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择形状", ex);
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

        if (disposing && _shape != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_shape) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _shape = null;
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
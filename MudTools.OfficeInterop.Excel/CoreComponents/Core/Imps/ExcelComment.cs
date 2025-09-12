//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Comment 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Comment 对象的安全访问和资源管理
/// </summary>
internal class ExcelComment : IExcelComment
{
    /// <summary>
    /// 底层的 COM Comment 对象
    /// </summary>
    private MsExcel.Comment _comment;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelComment 实例
    /// </summary>
    /// <param name="comment">底层的 COM Comment 对象</param>
    internal ExcelComment(MsExcel.Comment comment)
    {
        _comment = comment;
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
                // 释放子COM组件
                (_parent as ExcelRange)?.Dispose();
                (_shape as ExcelShape)?.Dispose();

                // 释放底层COM对象
                if (_comment != null)
                    Marshal.ReleaseComObject(_comment);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _comment = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);


    /// <summary>
    /// 获取或设置注释的文本内容
    /// </summary>
    public string Text(string? text = null, int? start = null, bool? overwrite = null)
    {
        return _comment.Text(text.ComArgsVal(), start.ComArgsVal(), overwrite.ComArgsVal());
    }

    /// <summary>
    /// 获取注释的作者
    /// </summary>
    public string Author => _comment?.Author?.ToString();

    /// <summary>
    /// 获取或设置注释是否可见
    /// </summary>
    public bool Visible
    {
        get => _comment != null && Convert.ToBoolean(_comment.Visible);
        set
        {
            if (_comment != null)
                _comment.Visible = value;
        }
    }

    /// <summary>
    /// 父级区域对象缓存
    /// </summary>
    private IExcelRange _parent;

    /// <summary>
    /// 获取注释所在的区域对象
    /// </summary>
    public IExcelRange Parent => _parent ??= new ExcelRange(_comment?.Parent as MsExcel.Range);

    /// <summary>
    /// 形状对象缓存
    /// </summary>
    private IExcelShape _shape;

    /// <summary>
    /// 获取注释的形状对象
    /// </summary>
    public IExcelShape Shape => _shape ??= new ExcelShape(_comment?.Shape);

    /// <summary>
    /// 删除注释
    /// </summary>
    public void Delete()
    {
        _comment?.Delete();
    }
}
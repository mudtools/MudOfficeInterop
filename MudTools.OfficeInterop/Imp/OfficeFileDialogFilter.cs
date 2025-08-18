//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;
/// <summary>
/// Office FileDialogFilter 对象的二次封装实现类
/// 实现 IOfficeFileDialogFilter 接口
/// </summary>
internal class OfficeFileDialogFilter : IOfficeFileDialogFilter
{
    internal MsCore.FileDialogFilter _filter;
    private bool _disposedValue = false;

    internal OfficeFileDialogFilter(MsCore.FileDialogFilter filter)
    {
        _filter = filter ?? throw new ArgumentNullException(nameof(filter));
    }

    #region 基础属性
    public object Parent => _filter.Parent;

    public string Description => _filter.Description;

    public string Extensions => _filter.Extensions;
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_filter != null)
                    Marshal.ReleaseComObject(_filter);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _filter = null;
        }

        _disposedValue = true;
    }

    ~OfficeFileDialogFilter()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}

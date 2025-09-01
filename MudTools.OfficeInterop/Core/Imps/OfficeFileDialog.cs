//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office FileDialog 对象的二次封装实现类
/// 实现 IFileDialog 接口
/// </summary>
internal class OfficeFileDialog : IOfficeFileDialog
{
    private MsCore.FileDialog _fileDialog;
    private bool _disposedValue = false;

    internal OfficeFileDialog(MsCore.FileDialog fileDialog)
    {
        _fileDialog = fileDialog ?? throw new ArgumentNullException(nameof(fileDialog));
    }

    #region 基础属性
    public object Parent => _fileDialog.Parent;

    public object Application => _fileDialog.Application;

    public MsoFileDialogType DialogType => (MsoFileDialogType)_fileDialog.DialogType;

    public string Title
    {
        get => _fileDialog.Title;
        set => _fileDialog.Title = value;
    }

    public string InitialFileName
    {
        get => _fileDialog.InitialFileName;
        set => _fileDialog.InitialFileName = value;
    }

    public MsoFileDialogView InitialView
    {
        get => (MsoFileDialogView)_fileDialog.InitialView;
        set => _fileDialog.InitialView = (MsCore.MsoFileDialogView)value;
    }

    public bool AllowMultiSelect
    {
        get => _fileDialog.AllowMultiSelect;
        set => _fileDialog.AllowMultiSelect = value;
    }

    public string ButtonName
    {
        get => _fileDialog.ButtonName;
        set => _fileDialog.ButtonName = value;
    }

    public int FilterIndex
    {
        get => _fileDialog.FilterIndex;
        set => _fileDialog.FilterIndex = value;
    }

    public IOfficeSelectedItems SelectedItems => new OfficeSelectedItems(_fileDialog.SelectedItems);
    #endregion


    #region 图表元素 (子对象)
    public IOfficeFileDialogFilters Filters => new OfficeFileDialogFilters(_fileDialog.Filters);
    #endregion

    #region 操作方法
    public int Show()
    {
        return _fileDialog.Show();
    }
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
                if (_fileDialog != null)
                    Marshal.ReleaseComObject(_fileDialog);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _fileDialog = null;
        }

        _disposedValue = true;
    }

    ~OfficeFileDialog()
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
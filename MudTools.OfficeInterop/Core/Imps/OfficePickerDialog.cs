//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.PickerDialog 的二次封装实现类。
/// 提供安全访问选取器对话框功能的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePickerDialog : IOfficePickerDialog
{
    private MsCore.PickerDialog _pickerDialog;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PickerDialog 对象。
    /// </summary>
    /// <param name="pickerDialog">原始的 COM PickerDialog 对象。</param>
    internal OfficePickerDialog(MsCore.PickerDialog pickerDialog)
    {
        _pickerDialog = pickerDialog ?? throw new ArgumentNullException(nameof(pickerDialog));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string DataHandlerId
    {
        get => _pickerDialog?.DataHandlerId ?? string.Empty;
        set
        {
            if (_pickerDialog != null)
                _pickerDialog.DataHandlerId = value;
        }
    }

    /// <inheritdoc/>
    public string Title
    {
        get => _pickerDialog?.Title ?? string.Empty;
        set
        {
            if (_pickerDialog != null)
                _pickerDialog.Title = value;
        }
    }

    /// <inheritdoc/>
    public IOfficePickerProperties Properties
    {
        get
        {
            if (_pickerDialog?.Properties != null)
            {
                try
                {
                    return new OfficePickerProperties(_pickerDialog.Properties);
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public bool IsPickerPropertiesCreated => _pickerDialog?.Properties != null;

    #endregion

    #region 方法实现
    /// <inheritdoc/>
    public IOfficePickerResults CreatePickerResults()
    {
        if (_pickerDialog == null)
            return null;

        try
        {
            var properties = _pickerDialog.CreatePickerResults();
            return properties != null ? new OfficePickerResults(properties) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficePickerResults Show(bool singleSelect = true, IOfficePickerResults? existingResults = null)
    {
        if (_pickerDialog == null)
            return null;

        try
        {
            MsCore.PickerResults objects = null;
            if (existingResults != null)
                objects = ((OfficePickerResults)existingResults)._pickerResults;
            var results = _pickerDialog.Show(singleSelect, objects);
            return results != null ? new OfficePickerResults(results) : null;
        }
        catch
        {
            return null;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pickerDialog != null)
        {
            try
            {
                Marshal.ReleaseComObject(_pickerDialog);
            }
            catch
            {
                // 忽略释放异常
            }
            _pickerDialog = null;
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
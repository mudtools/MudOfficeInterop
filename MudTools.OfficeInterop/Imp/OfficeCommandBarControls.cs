//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;

internal class OfficeCommandBarControls : IOfficeCommandBarControls
{
    private MsCore.CommandBarControls _controls;
    private bool _disposedValue;

    public int Count => _controls.Count;

    public IOfficeCommandBarControl this[int index] => new OfficeCommandBarControl(_controls[index]);

    internal OfficeCommandBarControls(MsCore.CommandBarControls controls)
    {
        _controls = controls ?? throw new ArgumentNullException(nameof(controls));
        _disposedValue = false;
    }

    public IOfficeCommandBarControl GetItem(object index)
    {
        try
        {
            var control = _controls[index];
            return control != null ? new OfficeCommandBarControl(control) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IOfficeCommandBarControl Add(MsoControlType type = MsoControlType.msoControlButton,
                                       object id = null, object parameter = null,
                                       object before = null, bool temporary = false)
    {
        try
        {
            object typeObj = (int)type;
            object idObj = id ?? Type.Missing;
            object paramObj = parameter ?? Type.Missing;
            object beforeObj = before ?? Type.Missing;

            var control = _controls.Add(typeObj, idObj, paramObj, beforeObj, temporary);
            return control != null ? new OfficeCommandBarControl(control) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加控件。", ex);
        }
    }


    public IOfficeCommandBar Parent => new OfficeCommandBar(_controls.Parent); // 假设OfficeCommandBar已实现

    public IEnumerator<IOfficeCommandBarControl> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _controls != null)
        {
            try
            {
                Marshal.ReleaseComObject(_controls);
            }
            catch { }
            _controls = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
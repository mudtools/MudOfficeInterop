//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;

internal class OfficeCommandBars : IOfficeCommandBars
{
    private MsCore.CommandBars _commandBars;
    private bool _disposedValue;

    public int Count => _commandBars.Count;

    public IOfficeCommandBar this[int index] => new OfficeCommandBar(_commandBars[index]);

    public IOfficeCommandBar this[string name] => new OfficeCommandBar(_commandBars[name]);

    internal OfficeCommandBars(MsCore.CommandBars commandBars)
    {
        _commandBars = commandBars ?? throw new ArgumentNullException(nameof(commandBars));
        _disposedValue = false;
    }

    public IOfficeCommandBar GetItem(object index)
    {
        try
        {
            var commandBar = _commandBars[index];
            return commandBar != null ? new OfficeCommandBar(commandBar) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IOfficeCommandBar Add(string name = null, MsoBarPosition position = MsoBarPosition.msoBarTop,
                                bool menuBar = false, bool temporary = false)
    {
        try
        {
            var commandBar = _commandBars.Add(name, (object)position, menuBar, temporary);
            return commandBar != null ? new OfficeCommandBar(commandBar) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加命令栏。", ex);
        }
    }

    public IOfficeCommandBarControl FindControl(int id)
    {
        try
        {
            var control = _commandBars.FindControl(Type.Missing, id, Type.Missing, Type.Missing);
            return control != null ? new OfficeCommandBarControl(control) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IOfficeCommandBarControls FindControls(int id)
    {
        try
        {
            var control = _commandBars.FindControls(Type.Missing, id, Type.Missing, Type.Missing);
            return control != null ? new OfficeCommandBarControls(control) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }


    public bool LargeButtons
    {
        get => _commandBars.LargeButtons;
        set => _commandBars.LargeButtons = value;
    }

    public MsoMenuAnimation MenuAnimationStyle
    {
        get => (MsoMenuAnimation)(int)_commandBars.MenuAnimationStyle;
        set => _commandBars.MenuAnimationStyle = (MsCore.MsoMenuAnimation)value;
    }

    public IEnumerator<IOfficeCommandBar> GetEnumerator()
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

        if (disposing && _commandBars != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_commandBars) > 0) { }
            }
            catch { }
            _commandBars = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private static readonly object Missing = Type.Missing;
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;

internal class OfficeCommandBar : IOfficeCommandBar
{
    internal MsCore.CommandBar _commandBar;
    private bool _disposedValue;

    public int Index => _commandBar.Index;

    public string Name
    {
        get => _commandBar.Name;
        set => _commandBar.Name = value;
    }

    public string NameLocal
    {
        get => _commandBar.NameLocal;
        set => _commandBar.NameLocal = value;
    }

    public bool Visible
    {
        get => _commandBar.Visible;
        set => _commandBar.Visible = value;
    }

    public MsoBarPosition Position
    {
        get => (MsoBarPosition)(int)_commandBar.Position;
        set => _commandBar.Position = (MsCore.MsoBarPosition)value;
    }

    public bool BuiltIn => _commandBar.BuiltIn;


    public int Height
    {
        get => _commandBar.Height;
        set => _commandBar.Height = value;
    }

    public int Width
    {
        get => _commandBar.Width;
        set => _commandBar.Width = value;
    }

    public int Left
    {
        get => _commandBar.Left;
        set => _commandBar.Left = value;
    }

    public int Top
    {
        get => _commandBar.Top;
        set => _commandBar.Top = value;
    }

    public bool Enabled
    {
        get => _commandBar.Enabled;
        set => _commandBar.Enabled = value;
    }

    public IOfficeCommandBarControls Controls => new OfficeCommandBarControls(_commandBar.Controls); // 伪代码实现

    public int Id => _commandBar.Id;


    internal OfficeCommandBar(MsCore.CommandBar commandBar)
    {
        _commandBar = commandBar ?? throw new ArgumentNullException(nameof(commandBar));
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _commandBar.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除命令栏。", ex);
        }
    }

    public void Reset()
    {
        try
        {
            _commandBar.Reset();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法重置命令栏。", ex);
        }
    }

    public void ShowPopup(int x = 0, int y = 0)
    {
        try
        {
            _commandBar.ShowPopup(x, y);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法显示命令栏弹出菜单。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _commandBar != null)
        {
            Marshal.ReleaseComObject(_commandBar);
            _commandBar = null;
        }

        _disposedValue = true;
    }

    public IOfficeCommandBarControl? FindControl(
        MsoControlType type = MsoControlType.msoControlButton,
        object? id = null, object? tag = null, object? visible = null)
    {
        try
        {
            object typeObj = (int)type;
            object idObj = id ?? Type.Missing;
            object tagObj = tag ?? Type.Missing;
            object visibleObj = visible ?? Type.Missing;

            var control = _commandBar.FindControl(typeObj, idObj, tagObj, visibleObj);
            return control != null ? new OfficeCommandBarControl(control) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
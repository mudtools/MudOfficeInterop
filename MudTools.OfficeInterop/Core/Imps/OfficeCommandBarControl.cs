//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

internal class OfficeCommandBarControl : IOfficeCommandBarControl
{
    protected MsCore.CommandBarControl _control;
    private bool _disposedValue;

    public int Index => _control.Index;

    public int Id
    {
        get => _control.Id;
    }

    public object Control => _control.Control;

    public MsoControlType Type => (MsoControlType)(int)_control.Type;


    public string Caption
    {
        get => _control.Caption;
        set => _control.Caption = value;
    }

    public bool Visible
    {
        get => _control.Visible;
        set => _control.Visible = value;
    }

    public bool Enabled
    {
        get => _control.Enabled;
        set => _control.Enabled = value;
    }

    public string Tag
    {
        get => _control.Tag;
        set => _control.Tag = value;
    }

    public string TooltipText
    {
        get => _control.TooltipText;
        set => _control.TooltipText = value;
    }

    public string HelpFile
    {
        get => _control.HelpFile;
        set => _control.HelpFile = value;
    }

    public int HelpContextId
    {
        get => _control.HelpContextId;
        set => _control.HelpContextId = value;
    }

    public int Left => _control.Left;

    public int Top => _control.Top;

    public int Width
    {
        get => _control.Width;
        set => _control.Width = value;
    }

    public int Height
    {
        get => _control.Height;
        set => _control.Height = value;
    }

    public string Parameter
    {
        get => _control.Parameter;
        set => _control.Parameter = value;
    }

    public IOfficeCommandBar Parent => new OfficeCommandBar(_control.Parent);

    internal OfficeCommandBarControl(MsCore.CommandBarControl control)
    {
        _control = control ?? throw new ArgumentNullException(nameof(control));
        _disposedValue = false;
    }

    public void Execute()
    {
        try
        {
            _control.Execute();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法执行控件。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _control.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除控件。", ex);
        }
    }

    public IOfficeCommandBarControl? Copy(
        IOfficeCommandBar? bar = null,
        IOfficeCommandBar? before = null)
    {
        try
        {
            object barObj = bar != null ? ((OfficeCommandBar)bar)._commandBar : System.Type.Missing;
            object beforeObj = before != null ? ((OfficeCommandBar)before)._commandBar : System.Type.Missing;

            var copiedControl = _control.Copy(barObj, beforeObj);
            return copiedControl != null ? new OfficeCommandBarControl(copiedControl) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制控件。", ex);
        }
    }

    public void Move(IOfficeCommandBar bar = null, IOfficeCommandBar before = null)
    {
        try
        {
            object barObj = bar != null ? ((OfficeCommandBar)bar)._commandBar : System.Type.Missing;
            object beforeObj = before != null ? ((OfficeCommandBar)before)._commandBar : System.Type.Missing;

            _control.Move(barObj, beforeObj);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动控件。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _control != null)
        {
            try
            {
                Marshal.ReleaseComObject(_control);
            }
            catch { }
            _control = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
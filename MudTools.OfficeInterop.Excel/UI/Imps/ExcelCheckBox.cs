//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelCheckBox : IExcelCheckBox
{
    private MsExcel.CheckBox _checkBox;
    private bool _disposedValue;

    public int Index
    {
        get => _checkBox.Index;
    }

    public string Name
    {
        get => _checkBox.Name;
        set => _checkBox.Name = value;
    }

    public double Left
    {
        get => _checkBox.Left;
        set => _checkBox.Left = value;
    }

    public double Top
    {
        get => _checkBox.Top;
        set => _checkBox.Top = value;
    }

    public double Width
    {
        get => _checkBox.Width;
        set => _checkBox.Width = value;
    }

    public double Height
    {
        get => _checkBox.Height;
        set => _checkBox.Height = value;
    }

    public XlFormControl Type => XlFormControl.xlCheckBox;

    public bool Visible
    {
        get => _checkBox.Visible;
        set => _checkBox.Visible = value;
    }

    public bool Locked
    {
        get => _checkBox.Locked;
        set => _checkBox.Locked = value;
    }

    public bool Enabled
    {
        get => _checkBox.Enabled;
        set => _checkBox.Enabled = value;
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_checkBox.Parent as MsExcel.Worksheet);

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_checkBox.Parent as MsExcel.Worksheet);


    public string LinkedCell
    {
        get => _checkBox.LinkedCell;
        set => _checkBox.LinkedCell = value;
    }

    public object Value
    {
        get => _checkBox.Value;
        set => _checkBox.Value = value;
    }

    public string Text
    {
        get => _checkBox.Text;
        set => _checkBox.Text = value;
    }

    public string Caption
    {
        get => _checkBox.Caption;
        set => _checkBox.Caption = value;
    }


    internal ExcelCheckBox(MsExcel.CheckBox checkBox)
    {
        _checkBox = checkBox ?? throw new ArgumentNullException(nameof(checkBox));
        _disposedValue = false;
    }

    public void Select(bool replace = true)
    {
        try
        {
            _checkBox.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择复选框。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _checkBox.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除复选框。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _checkBox.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制复选框。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _checkBox.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切复选框。", ex);
        }
    }

    public void Move(double left, double top)
    {
        try
        {
            _checkBox.Left = left;
            _checkBox.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动复选框。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        try
        {
            _checkBox.Width = width;
            _checkBox.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整复选框大小。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _checkBox != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_checkBox) > 0) { }
            }
            catch { }
            _checkBox = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
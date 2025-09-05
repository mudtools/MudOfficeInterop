//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelEditBox : IExcelEditBox
{
    private MsExcel.EditBox _editBox;
    private bool _disposedValue;

    public int Index
    {
        get => _editBox.Index;
    }

    public string Name
    {
        get => _editBox.Name;
        set => _editBox.Name = value;
    }

    public double Left
    {
        get => _editBox.Left;
        set => _editBox.Left = value;
    }

    public double Top
    {
        get => _editBox.Top;
        set => _editBox.Top = value;
    }

    public double Width
    {
        get => _editBox.Width;
        set => _editBox.Width = value;
    }

    public double Height
    {
        get => _editBox.Height;
        set => _editBox.Height = value;
    }

    public XlFormControl Type => XlFormControl.xlEditBox;

    public bool Visible
    {
        get => _editBox.Visible;
        set => _editBox.Visible = value;
    }

    public bool Locked
    {
        get => _editBox.Locked;
        set => _editBox.Locked = value;
    }

    public bool Enabled
    {
        get => _editBox.Enabled;
        set => _editBox.Enabled = value;
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_editBox.Parent as MsExcel.Worksheet);

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_editBox.Parent as MsExcel.Worksheet);

    public string Text
    {
        get => _editBox.Text;
        set => _editBox.Text = value;
    }

    public string Caption
    {
        get => _editBox.Caption;
        set => _editBox.Caption = value;
    }
    internal ExcelEditBox(MsExcel.EditBox editBox)
    {
        _editBox = editBox ?? throw new ArgumentNullException(nameof(editBox));
        _disposedValue = false;
    }

    public void Select(bool replace = true)
    {
        try
        {
            _editBox.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择编辑框。", ex);
        }
    }


    public void Delete()
    {
        try
        {
            _editBox.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除编辑框。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _editBox.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制编辑框。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _editBox.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切编辑框。", ex);
        }
    }

    public void Move(double left, double top)
    {
        try
        {
            _editBox.Left = left;
            _editBox.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动编辑框。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        try
        {
            _editBox.Width = width;
            _editBox.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整编辑框大小。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _editBox != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_editBox) > 0) { }
            }
            catch { }
            _editBox = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
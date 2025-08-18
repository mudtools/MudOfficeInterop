//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelDropDown : IExcelDropDown
{
    private MsExcel.DropDown _dropDown;
    private bool _disposedValue;

    public int Index
    {
        get => _dropDown.Index;
    }

    public string Name
    {
        get => _dropDown.Name;
        set => _dropDown.Name = value;
    }

    public double Left
    {
        get => _dropDown.Left;
        set => _dropDown.Left = value;
    }

    public double Top
    {
        get => _dropDown.Top;
        set => _dropDown.Top = value;
    }

    public double Width
    {
        get => _dropDown.Width;
        set => _dropDown.Width = value;
    }

    public double Height
    {
        get => _dropDown.Height;
        set => _dropDown.Height = value;
    }

    public XlFormControl Type => XlFormControl.xlDropDown;

    public bool Visible
    {
        get => _dropDown.Visible;
        set => _dropDown.Visible = value;
    }

    public bool Locked
    {
        get => _dropDown.Locked;
        set => _dropDown.Locked = value;
    }

    public bool Enabled
    {
        get => _dropDown.Enabled;
        set => _dropDown.Enabled = value;
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_dropDown.Parent as MsExcel.Worksheet);

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_dropDown.Parent as MsExcel.Worksheet);


    public string LinkedCell
    {
        get => _dropDown.LinkedCell;
        set => _dropDown.LinkedCell = value;
    }

    public string ListFillRange
    {
        get => _dropDown.ListFillRange;
        set => _dropDown.ListFillRange = value;
    }

    public int Value
    {
        get => _dropDown.Value;
        set => _dropDown.Value = value;
    }

    public string Text
    {
        get => _dropDown.Text;
        set => _dropDown.Text = value;
    }

    public int ListCount => _dropDown.ListCount;

    public int DropDownLines
    {
        get => _dropDown.DropDownLines;
        set => _dropDown.DropDownLines = value;
    }

    internal ExcelDropDown(MsExcel.DropDown dropDown)
    {
        _dropDown = dropDown ?? throw new ArgumentNullException(nameof(dropDown));
        _disposedValue = false;
    }

    public void Select(bool replace = true)
    {
        try
        {
            _dropDown.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择下拉框。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _dropDown.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除下拉框。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _dropDown.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制下拉框。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _dropDown.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切下拉框。", ex);
        }
    }

    public void Move(double left, double top)
    {
        try
        {
            _dropDown.Left = left;
            _dropDown.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动下拉框。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        try
        {
            _dropDown.Width = width;
            _dropDown.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整下拉框大小。", ex);
        }
    }

    public void RemoveAllItems()
    {
        try
        {
            _dropDown.RemoveAllItems();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除下拉框中的所有项目。", ex);
        }
    }
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dropDown != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_dropDown) > 0) { }
            }
            catch { }
            _dropDown = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
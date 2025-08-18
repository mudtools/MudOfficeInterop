//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelListBox : IExcelListBox
{
    private MsExcel.ListBox _listBox;
    private bool _disposedValue;

    public int Index
    {
        get => _listBox.Index;
    }

    public string Name
    {
        get => _listBox.Name;
        set => _listBox.Name = value;
    }

    public double Left
    {
        get => _listBox.Left;
        set => _listBox.Left = value;
    }

    public double Top
    {
        get => _listBox.Top;
        set => _listBox.Top = value;
    }

    public double Width
    {
        get => _listBox.Width;
        set => _listBox.Width = value;
    }

    public double Height
    {
        get => _listBox.Height;
        set => _listBox.Height = value;
    }

    public XlFormControl Type => XlFormControl.xlListBox;

    public bool Visible
    {
        get => _listBox.Visible;
        set => _listBox.Visible = value;
    }

    public bool Locked
    {
        get => _listBox.Locked;
        set => _listBox.Locked = value;
    }

    public bool Enabled
    {
        get => _listBox.Enabled;
        set => _listBox.Enabled = value;
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_listBox.Parent as MsExcel.Worksheet);

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_listBox.Parent as MsExcel.Worksheet);


    public string LinkedCell
    {
        get => _listBox.LinkedCell;
        set => _listBox.LinkedCell = value;
    }

    public string ListFillRange
    {
        get => _listBox.ListFillRange;
        set => _listBox.ListFillRange = value;
    }

    public int Value
    {
        get => _listBox.Value;
        set => _listBox.Value = value;
    }

    public string Text
    {
        get => _listBox.Value.ToString();
        set => _listBox.Value = int.Parse(value);
    }

    public int ListCount => _listBox.ListCount;

    internal ExcelListBox(MsExcel.ListBox listBox)
    {
        _listBox = listBox ?? throw new ArgumentNullException(nameof(listBox));
        _disposedValue = false;
    }

    public void Select(bool replace = true)
    {
        try
        {
            _listBox.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择列表框。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _listBox.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除列表框。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _listBox.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制列表框。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _listBox.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切列表框。", ex);
        }
    }

    public void Move(double left, double top)
    {
        try
        {
            _listBox.Left = left;
            _listBox.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动列表框。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        try
        {
            _listBox.Width = width;
            _listBox.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整列表框大小。", ex);
        }
    }


    public void SetItem(int index, string text)
    {
        if (index < 1 || index > ListCount)
            throw new ArgumentOutOfRangeException(nameof(index));

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("项目文本不能为空。", nameof(text));

        try
        {
            _listBox.List[index] = text;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法设置索引为 {index} 的项目。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listBox != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_listBox) > 0) { }
            }
            catch { }
            _listBox = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
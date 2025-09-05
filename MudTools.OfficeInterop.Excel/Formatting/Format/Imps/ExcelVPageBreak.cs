//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelVPageBreak : IExcelVPageBreak
{
    private MsExcel.VPageBreak _vPageBreak;
    private bool _disposedValue;


    public XlPageBreak Type
    {
        get => (XlPageBreak)(int)_vPageBreak.Type;
        set => _vPageBreak.Type = (MsExcel.XlPageBreak)value;
    }

    public IExcelRange Location
    {
        get => new ExcelRange(_vPageBreak.Location);
        set
        {
            if (value is ExcelRange excelRange)
            {
                _vPageBreak.Location = excelRange.InternalRange;
            }
        }
    }

    public int StartColumn => _vPageBreak.Location.Column;

    public int EndColumn => _vPageBreak.Location.Column + _vPageBreak.Location.Columns.Count - 1;

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_vPageBreak.Parent);

    public bool Enabled
    {
        get => _vPageBreak.Type != MsExcel.XlPageBreak.xlPageBreakNone;
        set => _vPageBreak.Type = value ? MsExcel.XlPageBreak.xlPageBreakManual : MsExcel.XlPageBreak.xlPageBreakNone;
    }

    internal ExcelVPageBreak(MsExcel.VPageBreak vPageBreak)
    {
        _vPageBreak = vPageBreak ?? throw new ArgumentNullException(nameof(vPageBreak));
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _vPageBreak.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除垂直分页符。", ex);
        }
    }

    public void MoveToColumn(int column)
    {
        if (column < 1)
            throw new ArgumentOutOfRangeException(nameof(column));

        try
        {
            var parent = _vPageBreak.Parent;
            var newRange = parent.Cells[1, column];
            _vPageBreak.Location = newRange as MsExcel.Range;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法将分页符移动到第 {column} 列。", ex);
        }
    }


    public IExcelRange GetPreviousColumnRange()
    {
        try
        {
            var worksheet = _vPageBreak.Parent;
            var startColumn = 1;
            var endColumn = StartColumn - 1;
            if (endColumn >= startColumn)
            {
                return new ExcelRange(worksheet.Range[worksheet.Cells[1, startColumn], worksheet.Cells[worksheet.UsedRange.Rows.Count, endColumn]]);
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取前一列的范围。", ex);
        }
    }

    public IExcelRange GetNextColumnRange()
    {
        try
        {
            var worksheet = _vPageBreak.Parent;
            var startColumn = EndColumn + 1;
            var endColumn = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            if (startColumn <= endColumn)
            {
                return new ExcelRange(worksheet.Range[worksheet.Cells[1, startColumn], worksheet.Cells[worksheet.UsedRange.Rows.Count, endColumn]]);
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取后一列的范围。", ex);
        }
    }

    public int GetAffectedColumnCount()
    {
        try
        {
            return EndColumn - StartColumn + 1;
        }
        catch (COMException)
        {
            return 0;
        }
    }

    public bool OverlapsWith(IExcelRange range)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        try
        {
            var rangeStartColumn = range.Column;
            var rangeEndColumn = rangeStartColumn + range.Columns.Count - 1;

            return !(EndColumn < rangeStartColumn || StartColumn > rangeEndColumn);
        }
        catch (COMException)
        {
            return false;
        }
    }

    public bool Validate()
    {
        try
        {
            return StartColumn >= 1 && EndColumn <= _vPageBreak.Parent.Columns.Count;
        }
        catch (COMException)
        {
            return false;
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _vPageBreak != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_vPageBreak) > 0) { }
            }
            catch { }
            _vPageBreak = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelHPageBreak : IExcelHPageBreak
{
    private MsExcel.HPageBreak? _hPageBreak;
    private bool _disposedValue;

    public XlPageBreak Type
    {
        get => _hPageBreak != null ? _hPageBreak.Type.EnumConvert(XlPageBreak.xlPageBreakNone) : XlPageBreak.xlPageBreakNone;
        set
        {
            if (_hPageBreak != null)
            {
                _hPageBreak.Type = value.EnumConvert(MsExcel.XlPageBreak.xlPageBreakNone);
            }
        }
    }

    public XlPageBreakExtent Extent
    {
        get => _hPageBreak != null ? _hPageBreak.Extent.EnumConvert(XlPageBreakExtent.xlPageBreakFull) : XlPageBreakExtent.xlPageBreakFull;
    }

    public IExcelRange? Location
    {
        get => _hPageBreak != null ? new ExcelRange(_hPageBreak.Location) : null;
        set
        {
            if (_hPageBreak != null && value is ExcelRange excelRange)
            {
                _hPageBreak.Location = excelRange.InternalRange;
            }
        }
    }

    public int StartRow => _hPageBreak != null ? _hPageBreak.Location.Row : 0;

    public int EndRow => _hPageBreak != null ? _hPageBreak.Location.Row + _hPageBreak.Location.Rows.Count - 1 : 0;


    public bool IsManual => _hPageBreak != null ? _hPageBreak.Type == MsExcel.XlPageBreak.xlPageBreakManual : false;

    public bool IsAutomatic => _hPageBreak != null ? _hPageBreak.Type == MsExcel.XlPageBreak.xlPageBreakAutomatic : false;

    public IExcelHPageBreaks? Parent => _hPageBreak != null ? new ExcelHPageBreaks(_hPageBreak.Parent.HPageBreaks) : null;

    public IExcelWorksheet? Worksheet => _hPageBreak != null ? new ExcelWorksheet(_hPageBreak.Parent) : null;

    internal ExcelHPageBreak(MsExcel.HPageBreak hPageBreak)
    {
        _hPageBreak = hPageBreak ?? throw new ArgumentNullException(nameof(hPageBreak));
        _disposedValue = false;
    }

    public void DragOff(XlDirection direction, int regionIndex)
    {
        try
        {
            _hPageBreak?.DragOff(direction.EnumConvert(MsExcel.XlDirection.xlDown), regionIndex);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法拖动分页符。", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法拖动分页符。", ex);
        }
    }
    public void Delete()
    {
        try
        {
            _hPageBreak?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除水平分页符。", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法删除水平分页符。", ex);
        }
    }

    public void MoveToRow(int row)
    {
        if (row < 1)
            throw new ArgumentOutOfRangeException(nameof(row));
        if (_hPageBreak == null)
            return;
        try
        {
            // 移动分页符到指定行（需要重新创建）
            var parent = _hPageBreak.Parent;
            var newRange = parent.Cells[row, 1];
            _hPageBreak.Location = newRange as MsExcel.Range;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法将分页符移动到第 {row} 行。", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法将分页符移动到第 {row} 行。", ex);
        }
    }


    public IExcelRange? GetPreviousPageRange()
    {
        if (_hPageBreak == null)
            return null;
        try
        {
            // 获取前一页的范围
            var worksheet = _hPageBreak.Parent;
            var startRow = 1;
            var endRow = StartRow - 1;
            if (endRow >= startRow)
            {
                return new ExcelRange(worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[endRow, worksheet.UsedRange.Columns.Count]]);
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取前一页的范围。", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法获取前一页的范围。", ex);
        }
    }

    public IExcelRange? GetNextPageRange()
    {
        if (_hPageBreak == null)
            return null;
        try
        {
            // 获取后一页的范围
            var worksheet = _hPageBreak.Parent;
            var startRow = EndRow + 1;
            var endRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            if (startRow <= endRow)
            {
                return new ExcelRange(worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[endRow, worksheet.UsedRange.Columns.Count]]);
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取后一页的范围。", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法获取后一页的范围。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _hPageBreak != null)
        {
            Marshal.ReleaseComObject(_hPageBreak);
            _hPageBreak = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
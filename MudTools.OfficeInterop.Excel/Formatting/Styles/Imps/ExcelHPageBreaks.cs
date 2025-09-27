//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelHPageBreaks : IExcelHPageBreaks
{
    private MsExcel.HPageBreaks? _hPageBreaks;
    private bool _disposedValue;
    private DisposableList _disposables = [];
    public int Count => _hPageBreaks != null ? _hPageBreaks.Count : 0;

    public IExcelHPageBreak? this[int index]
    {
        get
        {
            if (_hPageBreaks == null)
                throw new ObjectDisposedException(nameof(ExcelHPageBreaks));

            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));

            var hPageBreak = _hPageBreaks[index];
            var hPage = hPageBreak != null ? new ExcelHPageBreak(hPageBreak) : null;
            if (hPage != null)
            {
                _disposables.Add(hPage);
            }
            return hPage;
        }
    }

    internal ExcelHPageBreaks(MsExcel.HPageBreaks hPageBreaks)
    {
        _hPageBreaks = hPageBreaks ?? throw new ArgumentNullException(nameof(hPageBreaks));
        _disposedValue = false;
    }


    public IExcelHPageBreak? Add(IExcelRange before)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        if (before == null)
            throw new ArgumentNullException(nameof(before));

        try
        {
            MsExcel.Range? comRange = null;
            if (before is ExcelRange excelRange)
            {
                comRange = excelRange.InternalRange;
            }
            var hPageBreak = _hPageBreaks.Add(comRange);
            var hPage = hPageBreak != null ? new ExcelHPageBreak(hPageBreak) : null;
            if (hPage != null)
            {
                _disposables.Add(hPage);
            }
            return hPage;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加水平分页符。", ex);
        }
    }

    public IExcelHPageBreak? FindByRow(int row)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        if (row < 1)
            throw new ArgumentOutOfRangeException(nameof(row));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var pageBreak = this[i];
                if (pageBreak != null && pageBreak.StartRow <= row && pageBreak.EndRow >= row)
                {
                    return pageBreak;
                }
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法找到第 {row} 行的水平分页符。", ex);
        }
    }

    public void RemoveAt(int index)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        try
        {
            this[index].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除索引为 {index} 的水平分页符。", ex);
        }
    }

    public void RemoveByRow(int row)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        if (row < 1)
            throw new ArgumentOutOfRangeException(nameof(row));

        try
        {
            var pageBreak = FindByRow(row);
            pageBreak?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除第 {row} 行的水平分页符。", ex);
        }
    }

    public void Clear()
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        try
        {
            for (int i = Count; i >= 1; i--)
            {
                this[i].Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除所有水平分页符。", ex);
        }
    }

    public IExcelWorksheet? Parent => _hPageBreaks != null ? new ExcelWorksheet(_hPageBreaks.Parent as MsExcel.Worksheet) : null;

    public IExcelRange? Range => Parent?.UsedRange;

    public IEnumerable<IExcelHPageBreak> GetPageBreaksByType(XlPageBreak type)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        var result = new List<IExcelHPageBreak>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var pageBreak = this[i];
                if (pageBreak != null && pageBreak.Type == type)
                {
                    result.Add(pageBreak);
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取指定类型的水平分页符。", ex);
        }
        return result;
    }

    public IEnumerable<IExcelHPageBreak> GetPageBreaksInRange(int startRow, int endRow)
    {
        if (_hPageBreaks == null)
            throw new ObjectDisposedException(nameof(ExcelHPageBreaks));
        if (startRow < 1 || endRow < startRow)
            throw new ArgumentException("行号范围无效。");

        var result = new List<IExcelHPageBreak>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var pageBreak = this[i];
                if (pageBreak != null && pageBreak.StartRow >= startRow && pageBreak.StartRow <= endRow)
                {
                    result.Add(pageBreak);
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法获取指定范围内的水平分页符。", ex);
        }
        return result;
    }



    public IEnumerator<IExcelHPageBreak> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _hPageBreaks != null)
        {
            _disposables.Dispose();
            Marshal.ReleaseComObject(_hPageBreaks);
            _hPageBreaks = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
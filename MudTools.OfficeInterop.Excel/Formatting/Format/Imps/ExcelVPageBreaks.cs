//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelVPageBreaks : IExcelVPageBreaks
{
    private MsExcel.VPageBreaks _vPageBreaks;
    private bool _disposedValue;

    public int Count => _vPageBreaks.Count;

    public IExcelVPageBreak this[int index] => new ExcelVPageBreak(_vPageBreaks[index]);


    internal ExcelVPageBreaks(MsExcel.VPageBreaks vPageBreaks)
    {
        _vPageBreaks = vPageBreaks ?? throw new ArgumentNullException(nameof(vPageBreaks));
        _disposedValue = false;
    }

    public IExcelVPageBreak Add(IExcelRange before)
    {
        if (before == null)
            throw new ArgumentNullException(nameof(before));

        try
        {
            MsExcel.Range comRange = null;
            if (before is ExcelRange excelRange)
            {
                comRange = excelRange.InternalRange;
            }

            var vPageBreak = _vPageBreaks.Add(comRange);
            return vPageBreak != null ? new ExcelVPageBreak(vPageBreak) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加垂直分页符。", ex);
        }
    }


    public IExcelVPageBreak FindByRange(IExcelRange range)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var pageBreak = this[i];
                if (pageBreak.OverlapsWith(range))
                {
                    return pageBreak;
                }
            }
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IExcelVPageBreak FindByColumn(int column)
    {
        if (column < 1)
            throw new ArgumentOutOfRangeException(nameof(column));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var pageBreak = this[i];
                if (pageBreak.StartColumn <= column && pageBreak.EndColumn >= column)
                {
                    return pageBreak;
                }
            }
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }
    public void RemoveAt(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        try
        {
            this[index].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除索引为 {index} 的垂直分页符。", ex);
        }
    }

    public void RemoveByRange(IExcelRange range)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        try
        {
            var pageBreak = FindByRange(range);
            pageBreak?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移除指定范围的垂直分页符。", ex);
        }
    }

    public void RemoveByColumn(int column)
    {
        if (column < 1)
            throw new ArgumentOutOfRangeException(nameof(column));

        try
        {
            var pageBreak = FindByColumn(column);
            pageBreak?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除第 {column} 列的垂直分页符。", ex);
        }
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_vPageBreaks.Parent as MsExcel.Worksheet);

    public IExcelRange Range => Parent.UsedRange;




    public IEnumerator<IExcelVPageBreak> GetEnumerator()
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

        if (disposing && _vPageBreaks != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_vPageBreaks) > 0) { }
            }
            catch { }
            _vPageBreaks = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
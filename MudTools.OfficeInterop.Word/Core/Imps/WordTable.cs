//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 表格实现类
/// </summary>
internal class WordTable : IWordTable
{
    private readonly MsWord.Table _table;
    private bool _disposedValue;
    private IWordRange _range;

    public IWordApplication? Application => _table != null ? new WordApplication(_table.Application) : null;

    public int Rows => _table.Rows.Count;

    public int Columns => _table.Columns.Count;

    public object Parent => _table.Parent;

    public IWordRange Range
    {
        get
        {
            if (_range == null)
            {
                _range = new WordRange(_table.Range);
            }
            return _range;
        }
    }

    internal WordTable(MsWord.Table table)
    {
        _table = table ?? throw new ArgumentNullException(nameof(table));
        _disposedValue = false;
    }

    public IWordRange Cell(int row, int column)
    {
        if (row <= 0 || row > Rows || column <= 0 || column > Columns)
            throw new ArgumentOutOfRangeException("Row or column index out of range.");

        try
        {
            var cell = _table.Cell(row, column);
            return new WordRange(cell.Range);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get cell at row {row}, column {column}.", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _table.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete table.", ex);
        }
    }

    public void AutoFit()
    {
        try
        {
            _table.AutoFitBehavior(MsWord.WdAutoFitBehavior.wdAutoFitContent);
            _table.Columns.AutoFit();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to auto-fit table.", ex);
        }
    }

    public void SetBorders(bool enable = true)
    {
        try
        {
            if (enable)
            {
                _table.Borders.Enable = 1;
            }
            else
            {
                _table.Borders.Enable = 0;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set table borders.", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _range?.Dispose();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
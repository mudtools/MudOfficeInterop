//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档表格集合实现类
/// </summary>
internal class WordTables : IWordTables
{
    private MsWord.Tables? _tables;
    private DisposableList _disposables = [];
    private bool _disposedValue;

    public IWordApplication? Application => _tables != null ? new WordApplication(_tables.Application) : null;


    /// <summary>
    /// 获取表格数量
    /// </summary>
    public int Count => _tables?.Count ?? 0;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="tables">COM Tables 对象</param>
    internal WordTables(MsWord.Tables tables)
    {
        _tables = tables ?? throw new ArgumentNullException(nameof(tables));
        _disposedValue = false;
    }

    /// <inheritdoc/>
    public IWordTable? this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _tables == null) return null;
            try
            {
                var comTable = _tables[index];
                var table = comTable != null ? new WordTable(comTable) : null;
                if (table != null)
                    _disposables.Add(table);
                return table;
            }
            catch (COMException ex)
            {
                throw new ExcelOperationException("Failed to get table.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to get table.", ex);
            }
        }
    }

    /// <summary>
    /// 添加表格
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <param name="range">插入范围</param>
    /// <returns>表格对象</returns>
    public IWordTable? Add(IWordRange range, int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
            throw new ArgumentException("Rows and columns must be greater than zero.");
        if (range == null)
            throw new ArgumentNullException(nameof(range));
        if (_tables == null)
            return null;
        try
        {
            var comRange = GetComRange(range);
            var table = _tables.Add(comRange, rows, columns);
            return new WordTable(table);
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("Failed to add table.", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add table.", ex);
        }
    }

    /// <summary>
    /// 删除表格
    /// </summary>
    /// <param name="index">表格索引</param>
    public void Delete(int index)
    {
        try
        {
            var table = this[index];
            table?.Delete();
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException($"Failed to delete table at index {index}.", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete table at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>表格枚举器</returns>
    public IEnumerator<IWordTable> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 将 IWordRange 转换为 COM Range 对象
    /// </summary>
    /// <param name="range">IWordRange 对象</param>
    /// <returns>COM Range 对象</returns>
    private MsWord.Range GetComRange(IWordRange range)
    {
        return ((WordRange)range).InternalComObject;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tables != null)
        {
            _disposables?.Dispose();
            Marshal.ReleaseComObject(_tables);
            _tables = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

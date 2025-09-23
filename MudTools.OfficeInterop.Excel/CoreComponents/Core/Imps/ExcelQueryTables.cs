//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using System.Reflection;

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelQueryTables : IExcelQueryTables
{
    private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
    private DisposableList _disposables = [];
    internal MsExcel.QueryTables? _queryTables;
    private bool _disposedValue = false;

    internal ExcelQueryTables(MsExcel.QueryTables queryTables)
    {
        _queryTables = queryTables ?? throw new ArgumentNullException(nameof(queryTables));
    }

    public int Count => _queryTables != null ? _queryTables.Count : 0;

    public IExcelQueryTable? this[int index]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelQueryTables));
            if (_queryTables == null)
                return null;
            if (index < 1 || index > Count)
            {
                log.Error($"索引 {index} 超出范围");
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            try
            {
                var r = new ExcelQueryTable(_queryTables[index]);
                _disposables.Add(r);
                return r;
            }
            catch (COMException comEx)
            {
                log.Error($"获取第 {index} 个查询表失败: {comEx.Message}", comEx);
                throw;
            }
            catch (Exception ex)
            {
                log.Error($"获取第 {index} 个查询表失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    public object? Parent => _queryTables != null ? _queryTables.Parent : 0;
    public IExcelApplication? Application => _queryTables != null ? new ExcelApplication(_queryTables.Application) : null;

    public IExcelQueryTable? Add(object connection, IExcelRange destination, object sql = null)
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelQueryTables));
        if (destination == null) throw new ArgumentNullException(nameof(destination));

        if (_queryTables == null)
            return null;

        try
        {
            // 转换回原始 Range 对象
            var destRange = (destination is ExcelRange er) ? er._range : throw new ArgumentException("destination 必须是 ExcelRange 类型");

            MsExcel.QueryTable qt;
            if (sql != null)
                qt = _queryTables.Add(connection, destRange, sql);
            else
                qt = _queryTables.Add(connection, destRange);

            return new ExcelQueryTable(qt);
        }
        catch (COMException comEx)
        {
            log.Error($"添加查询表失败: {comEx.Message}", comEx);
            throw;
        }
        catch (Exception ex)
        {
            log.Error($"添加查询表失败: {ex.Message}", ex);
            throw;
        }
    }

    #region IEnumerable<IExcelQueryTable> Support

    public IEnumerator<IExcelQueryTable> GetEnumerator()
    {
        if (_disposedValue)
        {
            throw new ObjectDisposedException(nameof(ExcelQueryTables));
        }

        for (int i = 1; i <= _queryTables.Count; i++)
        {
            yield return new ExcelQueryTable((MsExcel.QueryTable)_queryTables[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable Support

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                _disposables?.Dispose();
                if (_queryTables != null)
                {
                    Marshal.ReleaseComObject(_queryTables);
                    _queryTables = null;
                }
            }
            catch (COMException comEx)
            {
                log.Error($"释放 QueryTables 时发生异常: {comEx.Message}", comEx);
            }
            catch (Exception ex)
            {
                log.Error($"释放 QueryTables 时发生异常: {ex.Message}", ex);
            }
        }

        _disposedValue = true;
    }

    ~ExcelQueryTables()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
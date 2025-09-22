//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelConnections : IExcelConnections
{
    private MsExcel.Connections _connections;
    private bool _disposedValue;

    public int Count => _connections.Count;

    public IExcelWorkbookConnection this[int index] => new ExcelWorkbookConnection(_connections[index]);

    public IExcelWorkbookConnection this[string name] => new ExcelWorkbookConnection(_connections[name]);

    internal ExcelConnections(MsExcel.Connections connections)
    {
        _connections = connections ?? throw new ArgumentNullException(nameof(connections));
        _disposedValue = false;
    }

    public IExcelWorkbookConnection GetItem(object index)
    {
        try
        {
            var connection = _connections.Item(index);
            return connection != null ? new ExcelWorkbookConnection(connection) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IExcelWorkbookConnection? Add(string name, string description, string connectionString,
                                       string? commandText = null, XlCmdType lCmdType = XlCmdType.xlCmdSql)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("连接名称不能为空。", nameof(name));

        if (string.IsNullOrEmpty(connectionString))
            throw new ArgumentException("连接字符串不能为空。", nameof(connectionString));

        try
        {
            var connection = _connections.Add(name, description, connectionString, commandText ?? string.Empty,
                                         lCmdType.EnumConvert(MsExcel.XlCmdType.xlCmdSql));
            return connection != null ? new ExcelWorkbookConnection(connection) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加连接: {name}", ex);
        }
    }

    public IExcelWorkbookConnection FindByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("连接名称不能为空。", nameof(name));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var connection = this[i];
                if (string.Equals(connection.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return connection;
                }
            }
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IEnumerable<IExcelWorkbookConnection> FindByType(XlConnectionType connectionType)
    {
        var result = new List<IExcelWorkbookConnection>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var connection = this[i];
                if (connection.Type == connectionType)
                {
                    result.Add(connection);
                }
            }
        }
        catch (COMException)
        {
            // 忽略异常，返回已找到的结果
        }
        return result;
    }

    public void Remove(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("连接名称不能为空。", nameof(name));

        try
        {
            var connection = FindByName(name);
            if (connection != null)
            {
                connection.Delete();
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除连接: {name}", ex);
        }
    }

    public void Clear()
    {
        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    this[i].Delete();
                }
                catch (COMException)
                {
                    // 继续删除其他连接
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除所有连接。", ex);
        }
    }

    public IExcelWorkbook Parent => new ExcelWorkbook(_connections.Parent as MsExcel.Workbook);

    public void RefreshAll()
    {
        try
        {
            this.Parent.RefreshAll();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法刷新所有连接。", ex);
        }
    }

    public IEnumerator<IExcelWorkbookConnection> GetEnumerator()
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

        if (disposing && _connections != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_connections) > 0) { }
            }
            catch { }
            _connections = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
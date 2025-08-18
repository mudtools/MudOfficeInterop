//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel RecentFiles 集合对象的二次封装实现类
/// 实现 IExcelRecentFiles 接口
/// </summary>
internal class ExcelRecentFiles : IExcelRecentFiles
{
    private MsExcel.RecentFiles _recentFiles;
    private bool _disposedValue = false;

    internal ExcelRecentFiles(MsExcel.RecentFiles recentFiles)
    {
        _recentFiles = recentFiles ?? throw new ArgumentNullException(nameof(recentFiles));
    }

    #region 基础属性
    public int Count => _recentFiles.Count;

    public IExcelRecentFile this[int index] => new ExcelRecentFile((MsExcel.RecentFile)_recentFiles[index]);

    public object Parent => _recentFiles.Parent;

    public IExcelApplication Application => new ExcelApplication(_recentFiles.Application);

    public int Maximum
    {
        get => _recentFiles.Maximum;
        set => _recentFiles.Maximum = value;
    }
    #endregion

    #region 查找和筛选
    public IExcelRecentFile[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IExcelRecentFile>();
        for (int i = 1; i <= Count; i++)
        {
            var file = this[i];
            if (string.Compare(file.Name, name, !matchCase) == 0)
            {
                results.Add(file);
            }
        }
        return results.ToArray();
    }

    public IExcelRecentFile[] FindByPath(string path, bool matchCase = false)
    {
        var results = new List<IExcelRecentFile>();
        for (int i = 1; i <= Count; i++)
        {
            var file = this[i];
            if (string.Compare(file.Path, path, !matchCase) == 0)
            {
                results.Add(file);
            }
        }
        return results.ToArray();
    }

    public IExcelRecentFile[] FindByExtension(string extension)
    {
        var results = new List<IExcelRecentFile>();
        for (int i = 1; i <= Count; i++)
        {
            var file = this[i];
            if (System.IO.Path.GetExtension(file.Name).Equals(extension, StringComparison.OrdinalIgnoreCase))
            {
                results.Add(file);
            }
        }
        return results.ToArray();
    }

    public IExcelRecentFile[] GetMostRecent(int count = 5)
    {
        // RecentFiles collection in Excel Interop is ordered by most recent first.
        var results = new List<IExcelRecentFile>();
        int itemsToTake = Math.Min(count, Count);
        for (int i = 1; i <= itemsToTake; i++)
        {
            results.Add(this[i]);
        }
        return results.ToArray();
    }

    public IExcelRecentFile[] GetLeastRecent(int count = 5)
    {
        var results = new List<IExcelRecentFile>();
        int startIndex = Math.Max(1, Count - count + 1);
        for (int i = Count; i >= startIndex; i--)
        {
            results.Add(this[i]);
        }
        return results.ToArray();
    }
    #endregion

    #region 操作方法
    public void Clear()
    {
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                Delete(i);
            }
            catch
            {
            }
        }
    }

    public void Delete(int index)
    {
        try
        {
            MsExcel.RecentFile fileToDelete = _recentFiles[index];
            fileToDelete?.Delete();
        }
        catch
        {
            // Handle error if index is invalid or deletion fails
        }
    }

    public void Delete(IExcelRecentFile file)
    {
        if (file is ExcelRecentFile excelFile)
        {
            try
            {
                excelFile._recentFile.Delete();
            }
            catch
            {
            }
        }
    }

    public void DeleteRange(int[] indices)
    {
        var sortedIndices = new List<int>(indices);
        sortedIndices.Sort((a, b) => b.CompareTo(a));
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }

    #endregion

    #region 高级功能
    public int OpenAll(bool readOnly = true)
    {
        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var file = this[i];
                file.Application.Workbooks.Open(file.Path, readOnly: readOnly);
                count++;
            }
            catch
            {
                // Log error or handle individual file open failure
            }
        }
        return count;
    }

    #endregion

    #region IEnumerable<IExcelRecentFile> Support
    public IEnumerator<IExcelRecentFile> GetEnumerator()
    {
        for (int i = 1; i <= _recentFiles.Count; i++)
        {
            yield return new ExcelRecentFile((MsExcel.RecentFile)_recentFiles[i]);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
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
                // 释放形状对象
                if (_recentFiles != null)
                    Marshal.ReleaseComObject(_recentFiles);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _recentFiles = null;
        }

        _disposedValue = true;
    }

    ~ExcelRecentFiles()
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

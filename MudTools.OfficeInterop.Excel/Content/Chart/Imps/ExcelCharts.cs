//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Charts 集合对象的二次封装实现类
/// 实现 IExcelCharts 接口
/// </summary>
internal class ExcelCharts : IExcelCharts
{
    private MsExcel.Charts _charts;
    private bool _disposedValue = false;

    internal ExcelCharts(MsExcel.Charts charts)
    {
        _charts = charts ?? throw new ArgumentNullException(nameof(charts));
    }

    #region 基础属性
    public int Count => _charts.Count;

    public IExcelChart this[int index]
    {
        get
        {
            if (_charts == null || index < 1 || index > Count)
                return null;

            try
            {
                var chartObject = _charts[index] as MsExcel.Chart;
                return chartObject != null ? new ExcelChart(chartObject) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public IExcelChart this[string name]
    {
        get
        {
            if (_charts == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var result = this.FindByName(name);
                if (result != null && result.Length > 0)
                    return result[0];
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    public object Parent => _charts.Parent;

    public IExcelApplication Application => new ExcelApplication(_charts.Application);
    #endregion

    #region 创建和添加
    public IExcelChart Add(object Before, object After, object Count)
    {
        var chartObj = _charts.Add(Before, After, Count);
        return new ExcelChart(chartObj);
    }
    #endregion

    #region 查找和筛选
    public IExcelChart[] FindByName(string name, bool matchCase = false)
    {
        List<IExcelChart> results = [];
        for (int i = 1; i <= Count; i++)
        {
            IExcelChart chart = this[i];
            if (string.Compare(chart.Name, name, !matchCase) == 0)
            {
                results.Add(chart);
            }
        }
        return [.. results];
    }

    public IExcelChart[] FindByType(MsoChartType chartType)
    {
        var results = new List<IExcelChart>();
        for (int i = 1; i <= Count; i++)
        {
            var chart = this[i];
            if (chart.ChartType == chartType)
            {
                results.Add(chart);
            }
        }
        return results.ToArray();
    }

    public IExcelChart[] GetProtectedCharts()
    {
        var results = new List<IExcelChart>();
        for (int i = 1; i <= Count; i++)
        {
            var chart = this[i];
            if (chart.IsProtected)
            {
                results.Add(chart);
            }
        }
        return results.ToArray();
    }

    public IExcelChart[] GetUnprotectedCharts()
    {
        var results = new List<IExcelChart>();
        for (int i = 1; i <= Count; i++)
        {
            var chart = this[i];
            if (!chart.IsProtected)
            {
                results.Add(chart);
            }
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
                ((MsExcel.Chart)_charts[i]).Delete();
            }
            catch { /* Handle error */ }
        }
    }

    public void Delete(int index)
    {
        try
        {
            ((MsExcel.Chart)_charts[index]).Delete();
        }
        catch { /* Handle error */ }
    }

    public void Delete(string name)
    {
        try
        {
            ((MsExcel.Chart)_charts[name]).Delete();
        }
        catch { /* Handle error */ }
    }

    public void Delete(IExcelChart chart)
    {
        if (chart is ExcelChart excelChart)
        {
            try
            {
                excelChart._chart.Delete();
            }
            catch { /* Handle error */ }
        }
    }

    /// <summary>
    /// 选择所有图表
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        _charts.Select(replace);
    }

    public void DeleteRange(int[] indices)
    {
        var sortedIndices = indices.OrderByDescending(i => i).ToArray();
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }

    public void Refresh()
    {
        for (int i = 1; i <= Count; i++)
        {
            this[i].Refresh();
        }
    }
    #endregion   

    #region 导出和导入
    public int ExportToFolder(string folderPath, string format = "png", string prefix = "chart_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            var chart = this[i];
            string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{i}.{format}");
            if (chart.ExportToImage(fileName, format))
            {
                count++;
            }
        }
        return count;
    }

    public byte[][] GetAllChartBytes()
    {
        var bytesList = new List<byte[]>();
        for (int i = 1; i <= Count; i++)
        {
            bytesList.Add(this[i].GetImageBytes());
        }
        return bytesList.ToArray();
    }
    #endregion    

    #region IEnumerable<IExcelChart> Support
    public IEnumerator<IExcelChart> GetEnumerator()
    {
        for (int i = 1; i <= _charts.Count; i++)
        {
            yield return new ExcelChart((MsExcel.Chart)_charts[i]);
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
                // 释放底层COM对象
                if (_charts != null)
                    Marshal.ReleaseComObject(_charts);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _charts = null;
        }
        _disposedValue = true;
    }

    ~ExcelCharts()
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

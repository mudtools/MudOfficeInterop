//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Trendlines 集合对象的二次封装实现类
/// 实现 IExcelTrendlines 接口
/// </summary>
internal class ExcelTrendlines : IExcelTrendlines
{
    private MsExcel.Trendlines _trendlines;
    private bool _disposedValue = false;

    internal ExcelTrendlines(MsExcel.Trendlines trendlines)
    {
        _trendlines = trendlines ?? throw new ArgumentNullException(nameof(trendlines));
    }

    #region 基础属性
    public int Count => _trendlines.Count;

    public IExcelTrendline this[int index]
    {
        get
        {
            return new ExcelTrendline(_trendlines.Item(index));
        }
    }

    public object Parent => _trendlines.Parent;

    public IExcelApplication Application => new ExcelApplication(_trendlines.Application);
    #endregion

    #region 创建和添加
    public IExcelTrendline Add(int type = 1, int order = 2, int period = 2, double forward = 0,
                               double backward = 0, double intercept = double.NaN, bool displayEquation = false,
                               bool displayRSquared = false, string name = "")
    {
        MsExcel.XlTrendlineType xlType = (MsExcel.XlTrendlineType)type;

        MsExcel.Trendline newTrendline = _trendlines.Add(
            Type: xlType,
            Order: order,
            Period: period,
            Forward: forward,
            Backward: backward,
            Intercept: double.IsNaN(intercept) ? System.Type.Missing : (object)intercept,
            DisplayEquation: displayEquation,
            DisplayRSquared: displayRSquared,
            Name: name
        );
        return new ExcelTrendline(newTrendline);
    }
    #endregion

    #region 查找和筛选
    public IExcelTrendline[] FindByType(int type)
    {
        List<IExcelTrendline> results = [];
        for (int i = 1; i <= Count; i++)
        {
            if (this[i].Type == type)
            {
                results.Add(this[i]);
            }
        }
        return results.ToArray();
    }

    public IExcelTrendline[] FindByName(string name, bool matchCase = false)
    {
        List<IExcelTrendline> results = [];
        for (int i = 1; i <= Count; i++)
        {
            if (string.Compare(this[i].Name, name, !matchCase) == 0)
            {
                results.Add(this[i]);
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
            MsExcel.Trendline trendlineToDelete = _trendlines.Item(index) as MsExcel.Trendline;
            trendlineToDelete?.Delete();
        }
        catch
        {
        }
    }

    public void Delete(IExcelTrendline trendline)
    {
        if (trendline is ExcelTrendline realTrendline)
        {
            try
            {
                realTrendline._trendline.Delete();
            }
            catch { }
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




    #region IEnumerable<IExcelTrendline> Support
    public IEnumerator<IExcelTrendline> GetEnumerator()
    {
        for (int i = 1; i <= _trendlines.Count; i++)
        {
            yield return new ExcelTrendline(_trendlines.Item(i));
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
                if (_trendlines != null)
                    Marshal.ReleaseComObject(_trendlines);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _trendlines = null;
        }
        _disposedValue = true;
    }

    ~ExcelTrendlines()
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

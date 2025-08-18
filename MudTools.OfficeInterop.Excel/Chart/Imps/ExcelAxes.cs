//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Axes 集合对象的二次封装实现类
/// 实现 IExcelAxes 接口
/// </summary>
internal class ExcelAxes : IExcelAxes
{
    private MsExcel.Axes _axes;
    private bool _disposedValue = false;

    internal ExcelAxes(MsExcel.Axes axes)
    {
        _axes = axes ?? throw new ArgumentNullException(nameof(axes));
    }

    #region 基础属性
    public int Count => _axes.Count;

    public IExcelAxis this[XlAxisType Type, XlAxisGroup AxisGroup = XlAxisGroup.xlPrimary]
    {
        get
        {
            MsExcel.Axis axis = _axes.Item((MsExcel.XlAxisType)Type, (MsExcel.XlAxisGroup)AxisGroup);
            return new ExcelAxis(axis);
        }
    }

    // 注意：通过索引获取 Axes 集合中的特定 Axis 在 Interop 中不直接支持，
    // 通常通过 Chart.Axes(type, group) 获取。这里为了接口一致性，提供一个简化实现。
    public IExcelAxis this[int index]
    {
        get
        {
            int i = 0;
            foreach (var axis in _axes)
            {
                if (i != index) continue;
                if (axis is MsExcel.Axis aObj && aObj != null)
                {
                    return new ExcelAxis(aObj);
                }
            }
            return null;
        }
    }

    public object Parent => _axes.Parent;

    public IExcelApplication Application => new ExcelApplication(_axes.Application);
    #endregion

    #region 查找和筛选
    public IExcelAxis GetAxis(XlAxisType type, XlAxisGroup group = XlAxisGroup.xlPrimary)
    {
        return new ExcelAxis(_axes.Item((MsExcel.XlAxisType)type, (MsExcel.XlAxisGroup)group)); // 假设 ExcelAxis 存在
    }
    #endregion

    #region IEnumerable/IEnumerator and IDisposable Support

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_axes != null)
                    Marshal.ReleaseComObject(_axes);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _axes = null;
        }
        _disposedValue = true;
    }

    ~ExcelAxes()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    public IEnumerator<IExcelAxis> GetEnumerator()
    {
        foreach (var axis in _axes)
        {
            if (axis is MsExcel.Axis aObj && aObj != null)
            {
                yield return new ExcelAxis(aObj);
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion
}

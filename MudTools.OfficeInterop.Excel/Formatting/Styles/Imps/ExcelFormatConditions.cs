//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel FormatConditions 集合对象的二次封装实现类
/// 实现 IExcelFormatConditions 接口
/// </summary>
internal class ExcelFormatConditions : IExcelFormatConditions
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelFormatConditions));
    private MsExcel.FormatConditions? _formatConditions;
    private bool _disposedValue = false;
    private DisposableList _disposables = [];

    internal ExcelFormatConditions(MsExcel.FormatConditions formatConditions)
    {
        _formatConditions = formatConditions ?? throw new ArgumentNullException(nameof(formatConditions));
    }

    #region 基础属性
    public int Count => _formatConditions != null ? _formatConditions.Count : 0;

    public IExcelFormatCondition? this[int index]
    {
        get
        {
            if (_formatConditions == null || index < 1 || index > Count)
                return null;

            try
            {
                var n = _formatConditions.Item(index) is MsExcel.FormatCondition name ? new ExcelFormatCondition(name) : null;
                if (n != null)
                {
                    _disposables.Add(n);
                }
                return n;
            }
            catch (COMException ex)
            {
                log.Error("获取指定索引的条件格式规则对象失败：" + ex.Message, ex);
                return null;
            }
            catch (Exception ex)
            {
                log.Error("获取指定索引的条件格式规则对象失败：" + ex.Message, ex);
                return null;
            }
        }
    }

    public object? Parent => _formatConditions?.Parent;

    public IExcelApplication? Application => _formatConditions != null ? new ExcelApplication(_formatConditions.Application) : null;
    #endregion

    #region 创建和添加
    public IExcelFormatCondition? Add(
        XlFormatConditionType type,
        XlFormatConditionOperator? @operator,
        string formula1 = "",
        string formula2 = "")
    {
        if (_formatConditions == null)
            return null;
        object oper = Type.Missing;
        if (@operator != null)
            oper = @operator;

        object formula1Obj = Type.Missing;
        if (!string.IsNullOrEmpty(formula1))
            formula1Obj = formula1;

        object formula2Obj = Type.Missing;
        if (!string.IsNullOrEmpty(formula2))
            formula2Obj = formula2;

        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            (MsExcel.XlFormatConditionType)type,
            oper,
            formula1Obj,
            formula2Obj
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddExpression(string formula)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlExpression,
            Type.Missing,
            formula,
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddColorScale(int colorScaleType)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlColorScale,
            Type.Missing,
            colorScaleType,
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddDatabar()
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlDatabar,
            Type.Missing,
            Type.Missing,
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddIconSetCondition(XlIconSet iconSet)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlIconSets,
            Type.Missing,
            iconSet.EnumConvert(MsExcel.XlIconSet.xlCustomSet),
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddUniqueValues(bool showUnique)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlUniqueValues,
            showUnique ? MsExcel.XlFormatConditionOperator.xlEqual : MsExcel.XlFormatConditionOperator.xlNotEqual,
            Type.Missing,
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    public IExcelFormatCondition? AddTop10(int rank, bool aboveAverage = true, bool percent = false)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlTop10,
            aboveAverage ? MsExcel.XlFormatConditionOperator.xlGreater : MsExcel.XlFormatConditionOperator.xlLess,
            rank,
            percent
        );
        return new ExcelFormatCondition(newCondition);
    }
    #endregion

    #region 查找和筛选
    public IExcelFormatCondition[] FindByType(XlFormatConditionType type)
    {
        if (_formatConditions == null)
            return [];
        var results = new List<IExcelFormatCondition>();
        for (int i = 1; i <= Count; i++)
        {
            var condition = this[i];
            if (condition != null && condition.Type == type)
            {
                results.Add(condition);
            }
        }
        return results.ToArray();
    }
    #endregion

    #region 操作方法

    /// <summary>
    /// 删除条件格式规则
    /// </summary>
    public void Delete()
    {
        _formatConditions?.Delete();
    }

    public void Delete(int index)
    {
        if (_formatConditions == null)
            return;
        var item = _formatConditions.Item(index);
        if (item != null)
        {
            ((MsExcel.FormatCondition)item).Delete();
        }
    }

    public void Delete(IExcelFormatCondition condition)
    {
        condition.Delete();
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

    #region IEnumerable<IExcelFormatCondition> Support
    public IEnumerator<IExcelFormatCondition> GetEnumerator()
    {
        if (_formatConditions == null)
            yield break;
        for (int i = 1; i <= _formatConditions.Count; i++)
        {
            yield return new ExcelFormatCondition((MsExcel.FormatCondition)_formatConditions.Item(i));
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
                _disposables.Dispose();
                // 释放形状对象
                if (_formatConditions != null)
                    Marshal.ReleaseComObject(_formatConditions);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _formatConditions = null;
        }

        _disposedValue = true;
    }

    ~ExcelFormatConditions()
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



namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// <see cref="IExcelCalculatedFields"/> 接口的内部实现类。
/// 负责包装 Microsoft.Office.Interop.Excel.CalculatedFields COM 对象，并管理其生命周期及子对象的生命周期。
/// </summary>
internal class ExcelCalculatedFields : IExcelCalculatedFields
{
    internal MsExcel.CalculatedFields? _calculatedFields;
    private DisposableList _disposables = new();
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelCalculatedFields));
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 <see cref="ExcelCalculatedFields"/> 类的新实例。
    /// </summary>
    /// <param name="calculatedFields">要包装的原始 COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="calculatedFields"/> 为 null 时抛出。</exception>
    internal ExcelCalculatedFields(MsExcel.CalculatedFields calculatedFields)
    {
        _calculatedFields = calculatedFields ?? throw new ArgumentNullException(nameof(calculatedFields));
    }

    /// <summary>
    /// 获取集合中的计算字段总数。
    /// </summary>
    public int Count => _calculatedFields?.Count ?? 0;

    /// <summary>
    /// 获取集合中指定索引或名称的计算字段。
    /// </summary>
    /// <param name="index">计算字段的索引（从1开始）或名称。</param>
    /// <returns>指定的 <see cref="IExcelPivotField"/> 对象。</returns>
    public IExcelPivotField? this[object index]
    {
        get
        {
            if (_calculatedFields == null) return null;
            try
            {
                var pivotField = _calculatedFields[index];
                var wrappedField = new ExcelPivotField(pivotField);
                _disposables.Add(wrappedField);
                return wrappedField;
            }
            catch (Exception ex)
            {
                log.Error($"获取计算字段 '{index}' 失败: {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    public object? Parent => _calculatedFields?.Parent;

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象。
    /// </summary>
    public IExcelApplication? Application => _calculatedFields != null ? new ExcelApplication(_calculatedFields.Application) : null;

    /// <summary>
    /// 在数据透视表中创建一个新的计算字段。
    /// </summary>
    /// <param name="name">新计算字段的名称。</param>
    /// <param name="formula">新计算字段的公式。</param>
    /// <param name="useStandardFormula">
    /// 如果为 true，则假定公式使用标准英语（美国）格式。
    /// 如果为 false，则假定公式采用本地化格式。
    /// </param>
    /// <returns>新创建的 <see cref="IExcelPivotField"/> 对象。</returns>
    public IExcelPivotField? Add(string name, string formula, bool useStandardFormula = true)
    {
        if (_calculatedFields == null) return null;
        if (string.IsNullOrEmpty(name)) throw new ArgumentException("计算字段名称不能为空。", nameof(name));
        if (string.IsNullOrEmpty(formula)) throw new ArgumentException("计算字段公式不能为空。", nameof(formula));

        try
        {
            // 调用 COM 对象的 Add 方法创建新的计算字段，该方法返回一个 PivotField 对象。
            MsExcel.PivotField newPivotField = _calculatedFields.Add(name, formula, useStandardFormula);
            var wrappedField = new ExcelPivotField(newPivotField);
            _disposables.Add(wrappedField);
            return wrappedField;
        }
        catch (Exception ex)
        {
            log.Error($"添加计算字段 '{name}' 失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 返回一个遍历集合中所有计算字段的枚举器。
    /// 枚举器通过索引 1 到 Count 来访问每个字段。
    /// </summary>
    /// <returns>一个 <see cref="IEnumerator{IExcelPivotField}"/> 对象。</returns>
    public IEnumerator<IExcelPivotField> GetEnumerator()
    {
        if (_calculatedFields == null)
            yield break;

        for (int i = 1; i <= _calculatedFields.Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 释放由 <see cref="ExcelCalculatedFields"/> 使用的资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _calculatedFields != null)
        {
            // 释放集合本身
            Marshal.ReleaseComObject(_calculatedFields);
            _calculatedFields = null;
            // 释放所有已创建的子 PivotField 对象
            _disposables.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保非托管资源被释放。
    /// </summary>
    ~ExcelCalculatedFields()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公共的 Dispose 方法，用于显式释放资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
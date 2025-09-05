//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel PivotFields 集合对象的二次封装实现类
/// 实现 IExcelPivotFields 接口
/// </summary>
internal class ExcelPivotFields : IExcelPivotFields
{
    private MsExcel.PivotFields _pivotFields;
    private bool _disposedValue = false;

    internal ExcelPivotFields(MsExcel.PivotFields pivotFields)
    {
        _pivotFields = pivotFields ?? throw new ArgumentNullException(nameof(pivotFields));
    }

    #region 基础属性
    public int Count => _pivotFields.Count;

    public IExcelPivotField this[int index]
    {
        get
        {
            if (_pivotFields == null || index < 1 || index > Count)
                return null;

            try
            {
                var pivotObject = _pivotFields.Item(index) as MsExcel.PivotField;
                return pivotObject != null ? new ExcelPivotField(pivotObject) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public IExcelPivotField this[string name]
    {
        get
        {
            if (_pivotFields == null || string.IsNullOrEmpty(name))
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

    public object Parent => _pivotFields.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotFields.Application);
    #endregion

    #region 查找和筛选
    public IExcelPivotField[] FindByName(string name, bool matchCase = false)
    {
        List<IExcelPivotField> results = [];
        for (int i = 1; i <= Count; i++)
        {
            IExcelPivotField field = this[i];
            if (string.Compare(field.Name, name, !matchCase) == 0)
            {
                results.Add(field);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotField[] FindByOrientation(XlPivotFieldOrientation orientation)
    {
        List<IExcelPivotField> results = [];
        for (int i = 1; i <= Count; i++)
        {
            IExcelPivotField field = this[i];
            if (field.Orientation == orientation)
            {
                results.Add(field);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotField[] FindByPosition(int position)
    {
        var results = new List<IExcelPivotField>();
        for (int i = 1; i <= Count; i++)
        {
            var field = this[i];
            if (field.Position == position)
            {
                results.Add(field);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotField[] GetCalculatedFields()
    {
        var results = new List<IExcelPivotField>();
        for (int i = 1; i <= Count; i++)
        {
            var field = this[i];
            if (field.IsCalculated)
            {
                results.Add(field);
            }
        }
        return results.ToArray();
    }
    #endregion

    #region 操作方法 
    public void Delete(int index)
    {
        try
        {
            var field = this[index];
            field.Orientation = (int)MsExcel.XlPivotFieldOrientation.xlHidden;
        }
        catch
        {

        }
    }

    public void Delete(string name)
    {
        try
        {
            var field = this[name];
            field.Orientation = (int)MsExcel.XlPivotFieldOrientation.xlHidden;
        }
        catch
        {
            // Handle error if name is invalid
        }
    }

    public void Delete(IExcelPivotField field)
    {
        if (field is ExcelPivotField excelField)
        {
            try
            {
                excelField._pivotField.Orientation = MsExcel.XlPivotFieldOrientation.xlHidden;
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


    #region IEnumerable<IExcelPivotField> Support
    public IEnumerator<IExcelPivotField> GetEnumerator()
    {
        for (int i = 1; i <= _pivotFields.Count; i++)
        {
            yield return new ExcelPivotField((MsExcel.PivotField)_pivotFields.Item(i));
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
                // 释放底层COM对象
                if (_pivotFields != null)
                    Marshal.ReleaseComObject(_pivotFields);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pivotFields = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotFields()
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

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ConditionValue 对象的二次封装实现类
/// 实现 IExcelConditionValue 接口
/// </summary>
internal class ExcelConditionValue : IExcelConditionValue
{
    private MsExcel.ConditionValue _conditionValue;
    private bool _disposedValue = false;

    internal ExcelConditionValue(MsExcel.ConditionValue conditionValue)
    {
        _conditionValue = conditionValue ?? throw new ArgumentNullException(nameof(conditionValue));
    }

    #region 基础属性
    public object Parent => _conditionValue.Parent;

    public IExcelApplication Application
    {
        get
        {
            var application = _conditionValue?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public XlConditionValueTypes Type
    {
        get => _conditionValue.Type.EnumConvert(XlConditionValueTypes.xlConditionValueNone);
    }

    public object Value
    {
        get => _conditionValue.Value;
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_conditionValue != null)
                Marshal.ReleaseComObject(_conditionValue);
            _conditionValue = null;
        }

        _disposedValue = true;
    }

    ~ExcelConditionValue()
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

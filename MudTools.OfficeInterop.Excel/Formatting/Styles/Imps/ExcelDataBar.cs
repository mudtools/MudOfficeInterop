//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Databar 对象的二次封装实现类
/// 实现 IExcelDataBar 接口
/// </summary>
internal class ExcelDataBar : IExcelDataBar
{
    private MsExcel.Databar? _databar;
    private bool _disposedValue = false;

    internal ExcelDataBar(MsExcel.Databar databar)
    {
        _databar = databar ?? throw new ArgumentNullException(nameof(databar));
    }

    #region 基础属性
    public object Parent => _databar.Parent;

    public IExcelApplication Application
    {
        get
        {
            var application = _databar?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public IExcelRange? AppliesTo
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelRange(_databar.AppliesTo);
        }
    }

    public IExcelConditionValue? MinPoint
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelConditionValue(_databar.MinPoint);
        }
    }

    public IExcelConditionValue? MaxPoint
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelConditionValue(_databar.MaxPoint);
        }
    }

    public int Direction
    {
        get => _databar != null ? _databar.Direction : 0;
        set
        {
            if (_databar != null)
                _databar.Direction = value;
        }
    }

    public XlDataBarFillType BarFillType
    {
        get => _databar.BarFillType.EnumConvert(XlDataBarFillType.xlDataBarFillSolid);
        set => _databar.BarFillType = value.EnumConvert(MsExcel.XlDataBarFillType.xlDataBarFillSolid);
    }

    public bool ShowBarOnly
    {
        get => _databar != null ? _databar.ShowValue : false;
        set
        {
            if (_databar != null)
                _databar.ShowValue = value;
        }
    }
    #endregion

    #region 格式设置
    public IExcelDataBarBorder? Borders
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelDataBarBorder(_databar.BarBorder);
        }
    }

    public IExcelFormatColor Color
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelFormatColor(_databar.Color);
        }
    }

    public IExcelNegativeBarFormat? NegativeBarFormat
    {
        get
        {
            if (_databar == null)
                return null;
            return new ExcelNegativeBarFormat(_databar.NegativeBarFormat);
        }
    }

    public string Formula
    {
        get
        {
            if (_databar == null)
                return string.Empty;
            return _databar.Formula;
        }
        set
        {
            if (_databar != null)
                _databar.Formula = value;
        }
    }

    public bool ShowValue
    {
        get => _databar != null ? _databar.ShowValue : false;
        set
        {
            if (_databar != null)
                _databar.ShowValue = value;
        }
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
                if (_databar != null)
                    Marshal.ReleaseComObject(_databar);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _databar = null;
        }

        _disposedValue = true;
    }

    ~ExcelDataBar()
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

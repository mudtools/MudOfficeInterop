//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Validation 对象的二次封装实现类
/// 实现 IExcelValidation 接口
/// </summary>
internal class ExcelValidation : IExcelValidation
{
    internal MsExcel.Validation? _validation;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelValidation));
    private bool _disposedValue = false;

    internal ExcelValidation(MsExcel.Validation validation)
    {
        _validation = validation ?? throw new ArgumentNullException(nameof(validation));
    }

    #region 基础属性
    public object? Parent => _validation?.Parent;

    public IExcelApplication? Application => _validation != null ? new ExcelApplication(_validation.Application) : null;

    public XlDVType Type
    {
        get => _validation != null ? _validation.Type.ObjectConvertEnum(XlDVType.xlValidateCustom) : XlDVType.xlValidateCustom;
        set => Modify(value, AlertStyle, Formula1, Formula2, Value);
    }

    public XlDVAlertStyle AlertStyle
    {
        get => _validation != null ? _validation.AlertStyle.ObjectConvertEnum(XlDVAlertStyle.xlValidAlertStop) : XlDVAlertStyle.xlValidAlertStop;
        set => Modify(Type, value, Formula1, Formula2, Value);
    }

    public string Formula1
    {
        get => _validation != null ? _validation.Formula1 : "";
        set => Modify(Type, AlertStyle, value, Formula2, Value);
    }

    public string Formula2
    {
        get => _validation != null ? _validation.Formula2 : "";
        set => Modify(Type, AlertStyle, Formula1, value, Value);
    }

    public bool Value
    {
        get => _validation != null ? _validation.Value : false;
        set => Modify(Type, AlertStyle, Formula1, Formula2, value);
    }

    public string InputTitle
    {
        get => _validation != null ? _validation.InputTitle : "";
        set
        {
            if (_validation != null)
            {
                _validation.InputTitle = value;
            }

        }
    }

    public string InputMessage
    {
        get => _validation != null ? _validation.InputMessage : "";
        set
        {
            if (_validation != null)
            {
                _validation.InputMessage = value;
            }

        }
    }

    public bool ShowError
    {
        get => _validation != null ? _validation.ShowError : false;
        set
        {
            if (_validation != null)
            {
                _validation.ShowError = value;
            }

        }
    }

    public bool ShowInput
    {
        get => _validation != null ? _validation.ShowInput : false;
        set
        {
            if (_validation != null)
            {
                _validation.ShowInput = value;
            }

        }
    }

    public string ErrorTitle
    {
        get => _validation != null ? _validation.ErrorTitle : "";
        set
        {
            if (_validation != null)
            {
                _validation.ErrorTitle = value;
            }

        }
    }

    public string ErrorMessage
    {
        get => _validation != null ? _validation.ErrorMessage : "";
        set
        {
            if (_validation != null)
            {
                _validation.ErrorMessage = value;
            }

        }
    }

    public bool IgnoreBlank
    {
        get => _validation != null ? _validation.IgnoreBlank : false;
        set
        {
            if (_validation != null)
            {
                _validation.IgnoreBlank = value;
            }

        }
    }

    public bool InCellDropdown
    {
        get => _validation != null ? _validation.InCellDropdown : false;
        set
        {
            if (_validation != null)
            {
                _validation.InCellDropdown = value;
            }

        }
    }
    #endregion

    #region 操作方法
    public void Delete()
    {
        try
        {
            _validation?.Delete();
        }
        catch (COMException cx)
        {
            log.Warn($"删除验证时发生异常:" + cx.Message, cx);
        }
        catch (Exception ex)
        {
            log.Warn($"删除验证时发生异常:" + ex.Message, ex);
        }
    }

    public void Add(XlDVType type, XlDVAlertStyle alertStyle, string formula1 = "", string formula2 = "", bool value = true)
    {
        try
        {
            _validation?.Add(
              type.EnumConvert(MsExcel.XlDVType.xlValidateCustom),
              alertStyle.EnumConvert(MsExcel.XlDVAlertStyle.xlValidAlertStop),
              value ? MsExcel.XlYesNoGuess.xlYes : MsExcel.XlYesNoGuess.xlNo,
              formula1,
              formula2
          );
        }
        catch (COMException cx)
        {
            log.Warn($"添加验证时发生异常:" + cx.Message, cx);
        }
        catch (Exception ex)
        {
            log.Warn($"添加验证时发生异常:" + ex.Message, ex);
        }
    }

    public void Modify(XlDVType type, XlDVAlertStyle alertStyle, string formula1 = "", string formula2 = "", bool value = true)
    {
        try
        {
            _validation?.Modify(
            type.EnumConvert(MsExcel.XlDVType.xlValidateCustom),
            alertStyle.EnumConvert(MsExcel.XlDVAlertStyle.xlValidAlertStop),
            value ? MsExcel.XlYesNoGuess.xlYes : MsExcel.XlYesNoGuess.xlNo,
            formula1,
            formula2
        );
        }
        catch (COMException cx)
        {
            log.Warn($"修改验证时发生异常:" + cx.Message, cx);
        }
        catch (Exception ex)
        {
            log.Warn($"修改验证时发生异常:" + ex.Message, ex);
        }
    }
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {

            if (_validation != null)
            {
                Marshal.ReleaseComObject(_validation);
                _validation = null;
            }
            _disposedValue = true;
        }
    }

    ~ExcelValidation()
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

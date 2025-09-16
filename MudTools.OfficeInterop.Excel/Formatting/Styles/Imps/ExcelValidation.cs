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
    internal MsExcel.Validation _validation;
    internal MsExcel.Range _appliedRange;
    private bool _disposedValue = false;

    internal ExcelValidation(MsExcel.Validation validation, MsExcel.Range appliedRange)
    {
        _validation = validation ?? throw new ArgumentNullException(nameof(validation));
        _appliedRange = appliedRange ?? throw new ArgumentNullException(nameof(appliedRange));
    }

    #region 基础属性
    public object Parent => _validation.Parent;

    public IExcelApplication Application => new ExcelApplication(_validation.Application);

    public int Type
    {
        get => (int)_validation.Type;
        set => Modify(value, AlertStyle, Formula1, Formula2, Value);
    }

    public int AlertStyle
    {
        get => (int)_validation.AlertStyle;
        set => Modify(Type, value, Formula1, Formula2, Value);
    }

    public string Formula1
    {
        get => _validation.Formula1;
        set => Modify(Type, AlertStyle, value, Formula2, Value);
    }

    public string Formula2
    {
        get => _validation.Formula2;
        set => Modify(Type, AlertStyle, Formula1, value, Value);
    }

    public bool Value
    {
        get => true;
        set => Modify(Type, AlertStyle, Formula1, Formula2, value);
    }

    public string InputTitle
    {
        get => _validation.InputTitle;
        set => _validation.InputTitle = value;
    }

    public string InputMessage
    {
        get => _validation.InputMessage;
        set => _validation.InputMessage = value;
    }

    public bool ShowError
    {
        get => _validation.ShowError;
        set => _validation.ShowError = value;
    }

    public string ErrorTitle
    {
        get => _validation.ErrorTitle;
        set => _validation.ErrorTitle = value;
    }

    public string ErrorMessage
    {
        get => _validation.ErrorMessage;
        set => _validation.ErrorMessage = value;
    }

    public bool IgnoreBlank
    {
        get => _validation.IgnoreBlank;
        set => _validation.IgnoreBlank = value;
    }

    public bool InCellDropdown
    {
        get => _validation.InCellDropdown;
        set => _validation.InCellDropdown = value;
    }
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        _appliedRange.Select();
    }

    public void Delete()
    {
        _validation.Delete();
    }

    public void Modify(int type, int alertStyle, string formula1 = "", string formula2 = "", bool value = true)
    {
        Delete();

        _appliedRange.Validation.Add(
            (MsExcel.XlDVType)type,
            (MsExcel.XlDVAlertStyle)alertStyle,
            value ? MsExcel.XlYesNoGuess.xlYes : MsExcel.XlYesNoGuess.xlNo,
            formula1,
            formula2
        );
        _validation = _appliedRange.Validation;
    }
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放托管状态(托管对象)
            }

            if (_validation != null)
            {
                try
                {
                    // Do not ReleaseComObject on _validation as it's a child of Range
                    // and managed by Excel's garbage collection with the Range.
                    // Only release the Range if this class owned it.
                    // Marshal.ReleaseComObject(_validation);
                }
                catch
                {
                    // 忽略释放过程中可能发生的异常
                }
                _validation = null;
            }
            _appliedRange = null;

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

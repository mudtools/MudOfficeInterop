//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel FormatCondition 对象的二次封装实现类
/// 实现 IExcelFormatCondition 接口
/// </summary>
internal class ExcelFormatCondition : IExcelFormatCondition
{
    internal MsExcel.FormatCondition? _formatCondition;
    private bool _disposedValue = false;

    internal ExcelFormatCondition(MsExcel.FormatCondition formatCondition)
    {
        _formatCondition = formatCondition ?? throw new ArgumentNullException(nameof(formatCondition));
    }

    #region 基础属性
    public object? Parent => _formatCondition?.Parent;

    public IExcelApplication? Application => _formatCondition != null ? new ExcelApplication(_formatCondition.Application) : null;

    public XlFormatConditionType Type => _formatCondition != null ? _formatCondition.Type.ObjectConvertEnum(XlFormatConditionType.xlNoBlanksCondition) : XlFormatConditionType.xlNoBlanksCondition;

    public XlFormatConditionOperator Operator
    {
        get => _formatCondition != null ? _formatCondition.Operator.ObjectConvertEnum(XlFormatConditionOperator.xlBetween) : XlFormatConditionOperator.xlBetween;
        set => Modify(Type, value, Formula1, Formula2);
    }

    public string Formula1
    {
        get => _formatCondition != null ? _formatCondition.Formula1 : "";
        set => Modify(Type, Operator, value, Formula2);
    }

    public string Formula2
    {
        get => _formatCondition != null ? _formatCondition.Formula2 : "";
        set => Modify(Type, Operator, Formula1, value);
    }

    public string Text
    {
        get => _formatCondition != null ? _formatCondition.Text : "";
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.Text = value;
            }
        }
    }

    public XlContainsOperator TextOperator
    {
        get => _formatCondition != null ? _formatCondition.TextOperator.EnumConvert(XlContainsOperator.xlContains) : XlContainsOperator.xlContains;
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.TextOperator = value.EnumConvert(MsExcel.XlContainsOperator.xlContains);
            }
        }
    }

    public XlPivotConditionScope ScopeType
    {
        get => _formatCondition != null ? _formatCondition.ScopeType.ObjectConvertEnum(XlPivotConditionScope.xlSelectionScope) : XlPivotConditionScope.xlSelectionScope;
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.ScopeType = value.EnumConvert(MsExcel.XlPivotConditionScope.xlSelectionScope);
            }
        }
    }

    public XlTimePeriods DateOperator
    {
        get => _formatCondition != null ? _formatCondition.DateOperator.ObjectConvertEnum(XlTimePeriods.xlYesterday) : XlTimePeriods.xlYesterday;
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.DateOperator = value.EnumConvert(MsExcel.XlTimePeriods.xlYesterday);
            }
        }
    }

    public bool PTCondition
    {
        get => _formatCondition != null && _formatCondition.PTCondition;
    }

    public bool StopIfTrue
    {
        get => _formatCondition != null && _formatCondition.StopIfTrue;
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.StopIfTrue = value;
            }
        }
    }

    public int Priority
    {
        get => _formatCondition != null ? _formatCondition.Priority : 0;
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.Priority = value;
            }
        }
    }
    #endregion

    #region 格式设置
    public IExcelFont? Font => _formatCondition != null ? new ExcelFont(_formatCondition.Font) : null;

    public IExcelInterior? Interior => _formatCondition != null ? new ExcelInterior(_formatCondition.Interior) : null;

    public IExcelBorders? Borders => _formatCondition != null ? new ExcelBorders(_formatCondition.Borders) : null;

    public IExcelRange? AppliesTo => _formatCondition != null ? new ExcelRange(_formatCondition.AppliesTo) : null;

    public string? NumberFormat
    {
        get => _formatCondition != null ? _formatCondition.NumberFormat.ToString() : "";
        set
        {
            if (_formatCondition != null)
            {
                _formatCondition.NumberFormat = value;
            }
        }
    }
    #endregion

    #region 操作方法 
    public void ModifyAppliesToRange(IExcelRange Range)
    {
        if (Range == null) throw new ArgumentNullException(nameof(Range));
        if (_formatCondition == null) throw new InvalidOperationException("无法修改格式条件。");

        _formatCondition?.ModifyAppliesToRange(((ExcelRange)Range).InternalRange);
    }

    public void SetFirstPriority()
    {
        _formatCondition?.SetFirstPriority();
    }

    public void SetLastPriority()
    {
        _formatCondition?.SetLastPriority();
    }

    public void Delete()
    {
        _formatCondition?.Delete();
    }

    public void Modify(XlFormatConditionType type, XlFormatConditionOperator @operator, string formula1, string formula2)
    {
        _formatCondition?.Modify(type.EnumConvert(MsExcel.XlFormatConditionType.xlNoBlanksCondition),
         @operator.EnumConvert(MsExcel.XlFormatConditionOperator.xlBetween),
          formula1, formula2);
    }

    public void ModifyExpression(string formula)
    {
        Modify(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlBetween, formula, "");
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_formatCondition != null)
                Marshal.ReleaseComObject(_formatCondition);
            _formatCondition = null;
        }

        _disposedValue = true;
    }

    ~ExcelFormatCondition()
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

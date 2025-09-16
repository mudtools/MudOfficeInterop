//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    internal MsExcel.FormatCondition _formatCondition;
    private bool _disposedValue = false;

    internal ExcelFormatCondition(MsExcel.FormatCondition formatCondition)
    {
        _formatCondition = formatCondition ?? throw new ArgumentNullException(nameof(formatCondition));
    }

    #region 基础属性
    public object Parent => _formatCondition.Parent;

    public IExcelApplication Application => new ExcelApplication(_formatCondition.Application);

    public int Type => (int)_formatCondition.Type;

    public int Operator
    {
        get => (int)_formatCondition.Operator;
        set => Modify(Type, value, Formula1, Formula2);
    }

    public string Formula1
    {
        get => _formatCondition.Formula1;
        set => Modify(Type, Operator, value, Formula2);
    }

    public string Formula2
    {
        get => _formatCondition.Formula2;
        set => Modify(Type, Operator, Formula1, value);
    }

    public string Text
    {
        get => _formatCondition.Text;
        set => _formatCondition.Text = value;
    }

    public int TextOperator
    {
        get => (int)_formatCondition.TextOperator;
        set => _formatCondition.TextOperator = (MsExcel.XlContainsOperator)value;
    }
    #endregion

    #region 格式设置
    public IExcelFont Font => new ExcelFont(_formatCondition.Font);

    public IExcelInterior Interior => new ExcelInterior(_formatCondition.Interior);

    public IExcelBorders Borders => new ExcelBorders(_formatCondition.Borders);

    public object NumberFormat
    {
        get => _formatCondition.NumberFormat;
        set => _formatCondition.NumberFormat = value;
    }
    #endregion

    #region 高级属性 (需要类型检查)
    public IExcelColorScaleCriteria ColorScaleCriteria
    {
        get
        {
            if (_formatCondition is MsExcel.ColorScale colorScale)
            {
                return new ExcelColorScaleCriteria(colorScale.ColorScaleCriteria);
            }
            return null;
        }
    }

    public IExcelDataBar DataBar
    {
        get
        {
            if (_formatCondition is MsExcel.Databar databar)
            {
                return new ExcelDataBar(databar);
            }
            return null;
        }
    }

    public IExcelIconSet IconSet
    {
        get
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                return new ExcelIconSet(iconSetCond.IconSet as MsExcel.IconSetCondition);
            }
            return null;
        }
    }

    public bool ShowIconOnly
    {
        get
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                return iconSetCond.ShowIconOnly;
            }
            return false;
        }
        set
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                iconSetCond.ShowIconOnly = value;
            }
        }
    }

    public bool ReverseOrder
    {
        get
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                return iconSetCond.ReverseOrder;
            }
            return false;
        }
        set
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                iconSetCond.ReverseOrder = value;
            }
        }
    }

    public bool PTCondition
    {
        get
        {
            if (_formatCondition is MsExcel.IconSetCondition iconSetCond)
            {
                return iconSetCond.PTCondition;
            }
            return false;
        }
    }
    #endregion

    #region 操作方法 
    public void Delete()
    {
        _formatCondition.Delete();
    }

    public void Modify(int type, int @operator, string formula1, string formula2)
    {
        _formatCondition.Modify((MsExcel.XlFormatConditionType)type, (MsExcel.XlFormatConditionOperator)@operator, formula1, formula2);
    }

    public void ModifyExpression(string formula)
    {
        Modify((int)MsExcel.XlFormatConditionType.xlExpression, (int)MsExcel.XlFormatConditionOperator.xlBetween, formula, "");
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
                if (_formatCondition != null)
                    Marshal.ReleaseComObject(_formatCondition);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
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

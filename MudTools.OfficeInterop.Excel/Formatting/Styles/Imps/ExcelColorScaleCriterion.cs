//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ColorScaleCriterion 对象的二次封装实现类
/// 实现 IExcelColorScaleCriterion 接口
/// </summary>
internal class ExcelColorScaleCriterion : IExcelColorScaleCriterion
{
    private MsExcel.ColorScaleCriterion _colorScaleCriterion;
    private readonly ExcelColorScaleCriteria _parentCriteria;
    private bool _disposedValue = false;

    internal ExcelColorScaleCriterion(ExcelColorScaleCriteria parentCriteria, MsExcel.ColorScaleCriterion colorScaleCriterion)
    {
        _parentCriteria = parentCriteria ?? throw new ArgumentNullException(nameof(parentCriteria));
        _colorScaleCriterion = colorScaleCriterion ?? throw new ArgumentNullException(nameof(colorScaleCriterion));
    }

    #region 基础属性

    public IExcelApplication Application => new ExcelApplication();

    public int Index => _colorScaleCriterion.Index;

    public XlConditionValueTypes Type
    {
        get => (XlConditionValueTypes)_colorScaleCriterion.Type;
        set => _colorScaleCriterion.Type = (MsExcel.XlConditionValueTypes)value;
    }

    public object Value
    {
        get => _colorScaleCriterion.Value;
        set => _colorScaleCriterion.Value = value;
    }

    public int Color
    {
        get
        {
            var formatColor = _colorScaleCriterion.FormatColor;
            return Convert.ToInt32(formatColor.Color);
        }
        set
        {
            var formatColor = _colorScaleCriterion.FormatColor;
            formatColor.Color = value;
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
                if (_colorScaleCriterion != null)
                    Marshal.ReleaseComObject(_colorScaleCriterion);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _colorScaleCriterion = null;
        }

        _disposedValue = true;
    }

    ~ExcelColorScaleCriterion()
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

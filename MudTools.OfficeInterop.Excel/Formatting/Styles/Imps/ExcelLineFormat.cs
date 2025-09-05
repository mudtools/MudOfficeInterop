//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel LineFormat 对象的二次封装实现类
/// </summary>
internal class ExcelLineFormat : IExcelLineFormat
{
    private MsExcel.LineFormat _lineFormat;
    private bool _disposedValue;

    internal ExcelLineFormat(MsExcel.LineFormat lineFormat)
    {
        _lineFormat = lineFormat;
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_lineFormat != null)
                    Marshal.ReleaseComObject(_lineFormat);
            }
            catch { }
            _lineFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public int Color
    {
        get => _lineFormat != null ? Convert.ToInt32(_lineFormat.ForeColor.RGB) : 0;
        set { if (_lineFormat != null) _lineFormat.ForeColor.RGB = value; }
    }

    public MsoLineDashStyle Style
    {
        get => _lineFormat != null ? (MsoLineDashStyle)_lineFormat.DashStyle : 0;
        set
        {
            if (_lineFormat != null)
                _lineFormat.DashStyle = (MsCore.MsoLineDashStyle)value;
        }
    }

    public float Weight
    {
        get => _lineFormat != null ? _lineFormat.Weight : 0;
        set { if (_lineFormat != null) _lineFormat.Weight = value; }
    }

    public bool Visible
    {
        get => _lineFormat != null && Convert.ToBoolean(_lineFormat.Visible);
        set
        {
            if (_lineFormat != null)
                _lineFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }
}
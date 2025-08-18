//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ThreeDFormat 对象的二次封装实现类
/// </summary>
internal class ExcelThreeDFormat : IExcelThreeDFormat
{
    private MsExcel.ThreeDFormat _threeDFormat;
    private bool _disposedValue;

    internal ExcelThreeDFormat(MsExcel.ThreeDFormat threeDFormat)
    {
        _threeDFormat = threeDFormat;
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_threeDFormat != null)
                    Marshal.ReleaseComObject(_threeDFormat);
            }
            catch { }
            _threeDFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public float Depth
    {
        get => _threeDFormat?.Depth ?? 0;
        set { if (_threeDFormat != null) _threeDFormat.Depth = value; }
    }

    public float RotationX
    {
        get => _threeDFormat?.RotationX ?? 0;
        set { if (_threeDFormat != null) _threeDFormat.RotationX = value; }
    }

    public float RotationY
    {
        get => _threeDFormat?.RotationY ?? 0;
        set { if (_threeDFormat != null) _threeDFormat.RotationY = value; }
    }

    public bool Perspective
    {
        get => _threeDFormat != null && Convert.ToBoolean(_threeDFormat.Perspective);
        set { if (_threeDFormat != null) _threeDFormat.Perspective = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse; }
    }

    public bool Visible
    {
        get => _threeDFormat != null && Convert.ToBoolean(_threeDFormat.Visible);
        set { if (_threeDFormat != null) _threeDFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse; }
    }
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel DataBarBorder 对象的二次封装实现类
/// 实现 IExcelDataBarBorder 接口
/// </summary>
internal class ExcelDataBarBorder : IExcelDataBarBorder
{
    private MsExcel.DataBarBorder _dataBarBorder;
    private bool _disposedValue = false;

    internal ExcelDataBarBorder(MsExcel.DataBarBorder dataBarBorder)
    {
        _dataBarBorder = dataBarBorder ?? throw new ArgumentNullException(nameof(dataBarBorder));
    }

    #region 基础属性
    public object Parent => _dataBarBorder.Parent;

    public IExcelApplication Application
    {
        get
        {
            var application = _dataBarBorder?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public XlDataBarBorderType Type
    {
        get => _dataBarBorder != null ? _dataBarBorder.Type.EnumConvert(XlDataBarBorderType.xlDataBarBorderNone) : XlDataBarBorderType.xlDataBarBorderNone;
        set
        {
            if (_dataBarBorder != null)
                _dataBarBorder.Type = value.EnumConvert(MsExcel.XlDataBarBorderType.xlDataBarBorderNone);
        }
    }

    public int Color
    {
        get => Convert.ToInt32(_dataBarBorder.Color);
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
                if (_dataBarBorder != null)
                    Marshal.ReleaseComObject(_dataBarBorder);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _dataBarBorder = null;
        }

        _disposedValue = true;
    }

    ~ExcelDataBarBorder()
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

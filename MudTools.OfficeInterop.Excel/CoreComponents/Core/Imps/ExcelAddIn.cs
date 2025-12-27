//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel AddIn 对象的二次封装实现类
/// 实现 IExcelAddIn 接口
/// </summary>
internal class ExcelAddIn : IExcelAddIn
{
    private MsExcel.AddIn _addIn;
    private bool _disposedValue = false;

    internal ExcelAddIn(MsExcel.AddIn addIn)
    {
        _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
    }

    #region 基础属性
    public string Name
    {
        get => _addIn.Name;
    }
    public string FullName => _addIn.FullName;

    public string Title => _addIn.Title;
    public string Subject => _addIn.Subject;
    public string Path => _addIn.Path;

    public object? Parent => _addIn.Parent;

    public IExcelApplication? Application => new ExcelApplication(_addIn.Application);

    public string ProgId => _addIn.progID;

    public string CLSID => _addIn.CLSID;

    public string Comments => _addIn.Comments;

    public string Author => _addIn.Author;

    public string Keywords => _addIn.Keywords;
    #endregion

    #region 状态属性
    public bool Installed
    {
        get => _addIn.Installed;
        set => _addIn.Installed = value;
    }

    public bool IsOpen => _addIn.IsOpen;
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_addIn != null)
                    Marshal.ReleaseComObject(_addIn);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _addIn = null;
        }
        _disposedValue = true;
    }

    ~ExcelAddIn()
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
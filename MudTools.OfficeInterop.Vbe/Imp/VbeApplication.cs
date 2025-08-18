//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe.Imp;
/// <summary>
/// VBE Application 对象的二次封装实现类
/// 实现 IVbeApplication 接口
/// </summary>
internal class VbeApplication : IVbeApplication
{
    private MsVb.VBE _vbe;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 VbeApplication 实例
    /// </summary>
    /// <param name="vbe">要封装的 Microsoft.Vbe.Interop.VBE 对象</param>
    internal VbeApplication(MsVb.VBE vbe)
    {
        _vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
    }

    #region 基础属性

    public string Version => _vbe.Version;

    #endregion

    #region 状态属性
    public bool Visible
    {
        get => _vbe.MainWindow.Visible;
        set => _vbe.MainWindow.Visible = value;
    }
    #endregion

    #region 核心对象集合和属性
    public IVbeVBProjects VBProjects => _vbe.VBProjects != null ? new VbeVBProjects(_vbe.VBProjects) : null;

    public object ActiveVBProject => _vbe.ActiveVBProject; // object placeholder

    public IVbeVBComponent ActiveVBComponent => _vbe.SelectedVBComponent != null ? new VbeVBComponent(_vbe.SelectedVBComponent) : null;

    public object ActiveCodePane => _vbe.ActiveCodePane; // object placeholder
    #endregion

    #region 环境和设置
    public vbext_WindowState WindowState
    {
        get => (vbext_WindowState)_vbe.MainWindow.WindowState;
        set => _vbe.MainWindow.WindowState = (MsVb.vbext_WindowState)value;
    }

    public int Left
    {
        get => _vbe.MainWindow.Left;
        set => _vbe.MainWindow.Left = value;
    }

    public int Top
    {
        get => _vbe.MainWindow.Top;
        set => _vbe.MainWindow.Top = value;
    }

    public int Width
    {
        get => _vbe.MainWindow.Width;
        set => _vbe.MainWindow.Width = value;
    }

    public int Height
    {
        get => _vbe.MainWindow.Height;
        set => _vbe.MainWindow.Height = value;
    }
    #endregion

    #region 操作方法  

    public void Quit()
    {
        _vbe.MainWindow.Close();
    }

    public void SaveAll()
    {
        System.Diagnostics.Debug.WriteLine("Saving all VBProjects.");
        try
        {
            for (int i = 1; i <= this.VBProjects.Count; i++)
            {
                this.VBProjects[i].Save();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving all projects: {ex.Message}");
        }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {

            _vbe = null; // Just nullify the reference

            _disposedValue = true;
        }
    }

    ~VbeApplication()
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

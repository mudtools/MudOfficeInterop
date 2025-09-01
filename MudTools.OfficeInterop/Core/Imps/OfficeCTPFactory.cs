//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office CTP (Custom Task Pane) Factory 对象的二次封装实现类
/// 实现 IOfficeCTPFactory 接口
/// </summary>
internal class OfficeCTPFactory : IOfficeCTPFactory
{
    private MsCore.ICTPFactory _ctpFactory;
    private object _application;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 OfficeCTPFactory 实例
    /// </summary>
    /// <param name="ctpFactory">要封装的 Microsoft.Office.Core.ICTPFactory 对象</param>
    /// <param name="application">关联的 Application 对象 (可选)</param>
    internal OfficeCTPFactory(MsCore.ICTPFactory ctpFactory, object application = null)
    {
        _ctpFactory = ctpFactory ?? throw new ArgumentNullException(nameof(ctpFactory));
        _application = application;
    }

    #region 基础属性
    public object Application
    {
        get
        {
            return _application;
        }
    }
    #endregion

    #region 创建和添加
    public IOfficeCustomTaskPane CreateCTP(string CTPAxID, string title, object CTPParentWindow)
    {
        try
        {

            object comCustomTaskPane = _ctpFactory.CreateCTP(CTPAxID, title, CTPParentWindow);

            // 将返回的 COM 对象包装成我们的接口实现
            if (comCustomTaskPane is MsCore.CustomTaskPane ctpObject)
            {
                OfficeCustomTaskPane wrappedPane = new(ctpObject)
                {
                    Visible = true
                };
                return wrappedPane;
            }
            else
            {
                return null;
            }
        }
        catch (COMException comEx)
        {
            System.Diagnostics.Debug.WriteLine($"COM Error creating CTP: {comEx.Message}");
            throw;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error creating CTP: {ex.Message}");
            throw;
        }
    }
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {

            //if (_ctpFactory != null)
            //{
            //    try
            //    {
            //        while (Marshal.ReleaseComObject(_ctpFactory) > 0) { }
            //    }
            //    catch { }
            //    _ctpFactory = null;
            //}
            //_ctpFactory = null;
            _disposedValue = true;
        }
    }

    ~OfficeCTPFactory()
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

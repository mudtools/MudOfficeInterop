//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.PowerPoint.OLEFormat 的封装实现类
/// </summary>
internal class PowerPointOLEFormat : IPowerPointOLEFormat
{
    #region 属性封装
    /// <summary>
    /// 获取 OLE 对象的程序标识符
    /// </summary>
    public string ProgID => _oleFormat.ProgID;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _oleFormat.Parent;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IPowerPointApplication Application => new PowerPointApplication(_oleFormat.Application);


    /// <summary>
    /// 获取 OLE 对象的原始格式
    /// </summary>
    public object Object => _oleFormat.Object;

    #endregion

    #region 构造函数与私有字段

    private MsPowerPoint.OLEFormat _oleFormat;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 PowerPointOLEFormat 实例
    /// </summary>
    /// <param name="oleFormat">原始 COM OLEFormat 对象</param>
    internal PowerPointOLEFormat(MsPowerPoint.OLEFormat oleFormat)
    {
        _oleFormat = oleFormat ?? throw new ArgumentNullException(nameof(oleFormat));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    /// <summary>
    /// 激活 OLE 对象以进行编辑
    /// </summary>
    public void Activate()
    {
        try
        {
            _oleFormat.Activate();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法激活 OLE 对象", ex);
        }
    }


    /// <summary>
    /// 转换 OLE 对象的类型
    /// </summary>
    /// <param name="classType">新的类类型</param>
    public void DoVerb(int classType)
    {
        try
        {
            _oleFormat.DoVerb(classType);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法转换 OLE 对象类型", ex);
        }
    }

    #endregion

    #region IDisposable 模式实现

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否显式调用 Dispose()</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _oleFormat != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_oleFormat) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _oleFormat = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 显式释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
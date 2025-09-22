//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel SmartTag 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.SmartTag 对象的安全访问和资源管理
/// </summary>
internal class ExcelSmartTag : IExcelSmartTag
{
    /// <summary>
    /// 底层的 COM SmartTag 对象
    /// </summary>
    private MsExcel.SmartTag _smartTag;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelSmartTag 实例
    /// </summary>
    /// <param name="smartTag">底层的 COM SmartTag 对象</param>
    internal ExcelSmartTag(MsExcel.SmartTag smartTag)
    {
        _smartTag = smartTag;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放子COM组件
                (_range as ExcelRange)?.Dispose();
                (_smartTagActions as ExcelSmartTagActions)?.Dispose();

                // 释放底层COM对象
                if (_smartTag != null)
                    Marshal.ReleaseComObject(_smartTag);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _smartTag = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取智能标记的名称
    /// </summary>
    public string? Name => _smartTag?.Name;


    /// <summary>
    /// 获取智能标记的XML字符串
    /// </summary>
    public string? XML => _smartTag?.XML;

    /// <summary>
    /// 区域对象缓存
    /// </summary>
    private IExcelRange? _range;

    /// <summary>
    /// 获取智能标记所在的区域对象
    /// </summary>
    public IExcelRange? Range
    {
        get
        {
            _range ??= new ExcelRange(_smartTag?.Range);
            return _range;
        }
    }

    /// <summary>
    /// 智能标记动作集合缓存
    /// </summary>
    private IExcelSmartTagActions? _smartTagActions;

    /// <summary>
    /// 获取智能标记的动作集合
    /// </summary>
    public IExcelSmartTagActions? SmartTagActions
    {
        get
        {
            _smartTagActions ??= new ExcelSmartTagActions(_smartTag.SmartTagActions);
            return _smartTagActions;
        }
    }

    /// <summary>
    /// 删除智能标记
    /// </summary>
    public void Delete()
    {
        _smartTag?.Delete();
    }
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


using log4net;

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel SmartTagActions 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.SmartTagActions 对象的安全访问和资源管理
/// </summary>
internal class ExcelSmartTagActions : IExcelSmartTagActions
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSmartTagActions));
    /// <summary>
    /// 底层的 COM SmartTagActions 集合对象
    /// </summary>
    private MsExcel.SmartTagActions _smartTagActions;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelSmartTagActions 实例
    /// </summary>
    /// <param name="smartTagActions">底层的 COM SmartTagActions 集合对象</param>
    internal ExcelSmartTagActions(MsExcel.SmartTagActions smartTagActions)
    {
        _smartTagActions = smartTagActions;
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
                // 释放所有子智能标记动作对象
                for (int i = 1; i <= Count; i++)
                {
                    var action = this[i] as ExcelSmartTagAction;
                    action?.Dispose();
                }

                // 释放底层COM对象
                if (_smartTagActions != null)
                    Marshal.ReleaseComObject(_smartTagActions);
            }
            catch (Exception ex)
            {
                log.Warn("释放ExcelSmartTagActions资源时发生异常", ex);
                // 忽略释放过程中的异常
            }
            _smartTagActions = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取智能标记动作集合中的动作数量
    /// </summary>
    public int Count => _smartTagActions?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的智能标记动作对象
    /// </summary>
    /// <param name="index">智能标记动作索引（从1开始）</param>
    /// <returns>智能标记动作对象</returns>
    public IExcelSmartTagAction this[int index]
    {
        get
        {
            if (_smartTagActions == null || index < 1 || index > Count)
                return null;

            try
            {
                var action = _smartTagActions[index];
                return action != null ? new ExcelSmartTagAction(action) : null;
            }
            catch (Exception ex)
            {
                log.Warn($"获取索引为 {index} 的智能标记动作时发生异常", ex);
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的智能标记动作对象
    /// </summary>
    /// <param name="name">智能标记动作名称</param>
    /// <returns>智能标记动作对象</returns>
    public IExcelSmartTagAction this[string name]
    {
        get
        {
            if (_smartTagActions == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var action = _smartTagActions[name];
                return action != null ? new ExcelSmartTagAction(action) : null;
            }
            catch (Exception ex)
            {
                log.Warn($"获取名称为 '{name}' 的智能标记动作时发生异常", ex);
                return null;
            }
        }
    }


    public IEnumerator<IExcelSmartTagAction> GetEnumerator()
    {
        for (int i = 0; i < _smartTagActions.Count; i++)
        {
            yield return new ExcelSmartTagAction(_smartTagActions[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
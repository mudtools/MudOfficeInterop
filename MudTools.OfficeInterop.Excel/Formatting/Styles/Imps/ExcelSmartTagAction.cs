//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel SmartTagAction 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.SmartTagAction 对象的安全访问和资源管理
/// </summary>
internal class ExcelSmartTagAction : IExcelSmartTagAction
{
    /// <summary>
    /// 底层的 COM SmartTagAction 对象
    /// </summary>
    private MsExcel.SmartTagAction _smartTagAction;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelSmartTagAction 实例
    /// </summary>
    /// <param name="smartTagAction">底层的 COM SmartTagAction 对象</param>
    internal ExcelSmartTagAction(MsExcel.SmartTagAction smartTagAction)
    {
        _smartTagAction = smartTagAction;
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
            // 释放父级COM组件
            (_parent as ExcelSmartTag)?.Dispose();

            // 释放底层COM对象
            if (_smartTagAction != null)
                Marshal.ReleaseComObject(_smartTagAction);
            _smartTagAction = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取智能标记动作的名称
    /// </summary>
    public string Name => _smartTagAction?.Name;

    /// <summary>
    /// 获取智能标记动作的显示文本
    /// </summary>
    public string TextboxText => _smartTagAction?.TextboxText;

    /// <summary>
    /// 获取智能标记动作的类型
    /// </summary>
    public XlSmartTagControlType Type => _smartTagAction != null ? _smartTagAction.Type.EnumConvert(XlSmartTagControlType.xlSmartTagControlHelp) : XlSmartTagControlType.xlSmartTagControlHelp;

    /// <summary>
    /// 父级智能标记对象缓存
    /// </summary>
    private IExcelSmartTag? _parent;

    /// <summary>
    /// 获取智能标记动作所在的智能标记对象
    /// </summary>
    public IExcelSmartTag? Parent
    {
        get
        {
            if (_parent == null)
            {
                if (_smartTagAction.Parent is MsExcel.SmartTag smartTag)
                    _parent = new ExcelSmartTag(smartTag);
            }
            return _parent;
        }

    }

    /// <summary>
    /// 执行该智能标记动作
    /// </summary>
    public void Execute()
    {
        _smartTagAction?.Execute();
    }
}
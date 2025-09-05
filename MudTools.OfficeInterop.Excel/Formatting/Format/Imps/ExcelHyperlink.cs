//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Hyperlink 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Hyperlink 对象的安全访问和资源管理
/// </summary>
internal class ExcelHyperlink : IExcelHyperlink
{
    /// <summary>
    /// 底层的 COM Hyperlink 对象
    /// </summary>
    private MsExcel.Hyperlink _hyperlink;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelHyperlink 实例
    /// </summary>
    /// <param name="hyperlink">底层的 COM Hyperlink 对象</param>
    internal ExcelHyperlink(MsExcel.Hyperlink hyperlink)
    {
        _hyperlink = hyperlink;
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

                // 释放底层COM对象
                if (_hyperlink != null)
                    Marshal.ReleaseComObject(_hyperlink);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _hyperlink = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取超链接的名称
    /// </summary>
    public string Name => _hyperlink?.Name?.ToString();

    /// <summary>
    /// 获取或设置超链接的目标地址
    /// </summary>
    public string Address
    {
        get => _hyperlink?.Address?.ToString();
        set
        {
            if (_hyperlink != null && value != null)
                _hyperlink.Address = value;
        }
    }

    /// <summary>
    /// 获取或设置超链接的子地址
    /// </summary>
    public string SubAddress
    {
        get => _hyperlink?.SubAddress?.ToString();
        set
        {
            if (_hyperlink != null && value != null)
                _hyperlink.SubAddress = value;
        }
    }

    /// <summary>
    /// 获取或设置鼠标悬停时显示的提示文本
    /// </summary>
    public string ScreenTip
    {
        get => _hyperlink?.ScreenTip?.ToString();
        set
        {
            if (_hyperlink != null && value != null)
                _hyperlink.ScreenTip = value;
        }
    }

    /// <summary>
    /// 获取或设置要显示的文本
    /// </summary>
    public string TextToDisplay
    {
        get => _hyperlink?.TextToDisplay?.ToString();
        set
        {
            if (_hyperlink != null && value != null)
                _hyperlink.TextToDisplay = value;
        }
    }

    /// <summary>
    /// 区域对象缓存
    /// </summary>
    private IExcelRange _range;

    /// <summary>
    /// 获取超链接所在的区域对象
    /// </summary>
    public IExcelRange Range => _range ?? (_range = new ExcelRange(_hyperlink?.Range));

    /// <summary>
    /// 获取超链接的类型
    /// </summary>
    public int Type => _hyperlink != null ? Convert.ToInt32(_hyperlink.Type) : 0;

    /// <summary>
    /// 删除超链接
    /// </summary>
    public void Delete()
    {
        _hyperlink?.Delete();
    }

    /// <summary>
    /// 跟随超链接（打开链接）
    /// </summary>
    /// <param name="newWindow">是否在新窗口中打开</param>
    /// <param name="addHistory">是否添加到历史记录</param>
    /// <param name="extraInfo">额外信息</param>
    public void Follow(bool newWindow, bool addHistory, object extraInfo)
    {
        _hyperlink?.Follow(newWindow, addHistory, extraInfo);
    }
}
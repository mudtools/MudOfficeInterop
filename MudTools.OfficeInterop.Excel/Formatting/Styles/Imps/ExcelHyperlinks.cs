//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Hyperlinks 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Hyperlinks 对象的安全访问和资源管理
/// </summary>
internal class ExcelHyperlinks : IExcelHyperlinks
{
    /// <summary>
    /// 底层的 COM Hyperlinks 集合对象
    /// </summary>
    private MsExcel.Hyperlinks? _hyperlinks;

    private DisposableList _disposables = [];

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelHyperlinks 实例
    /// </summary>
    /// <param name="hyperlinks">底层的 COM Hyperlinks 集合对象</param>
    internal ExcelHyperlinks(MsExcel.Hyperlinks hyperlinks)
    {
        _hyperlinks = hyperlinks;
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
            _disposables.Dispose();

            // 释放底层COM对象
            if (_hyperlinks != null)
                Marshal.ReleaseComObject(_hyperlinks);
            _hyperlinks = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取超链接集合中的超链接数量
    /// </summary>
    public int Count => _hyperlinks?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的超链接对象
    /// </summary>
    /// <param name="index">超链接索引（从1开始）</param>
    /// <returns>超链接对象</returns>
    public IExcelHyperlink? this[int index]
    {
        get
        {
            if (_hyperlinks == null || index < 1 || index > Count)
                return null;

            var hyperlink = _hyperlinks[index];
            var link = hyperlink != null ? new ExcelHyperlink(hyperlink) : null;
            if (link != null)
                _disposables.Add(link);

            return link;
        }
    }

    /// <summary>
    /// 向集合中添加新的超链接
    /// </summary>
    /// <param name="anchor">超链接的锚点区域</param>
    /// <param name="address">链接地址（如网页URL或文件路径）</param>
    /// <param name="subAddress">子地址（如工作表名称或单元格引用）</param>
    /// <param name="screenTip">鼠标悬停时显示的提示文本</param>
    /// <param name="textToDisplay">要显示的文本</param>
    /// <returns>新创建的超链接对象</returns>
    public IExcelHyperlink? Add(IExcelRange anchor, string address, string? subAddress = null, string? screenTip = null, string? textToDisplay = null)
    {
        if (_hyperlinks == null || anchor == null)
            return null;

        object subAddressObj = Type.Missing;
        if (subAddress != null)
            subAddressObj = subAddress;

        object screenTipObj = Type.Missing;
        if (screenTip != null)
            screenTipObj = screenTip;

        object textToDisplayObj = Type.Missing;
        if (textToDisplay != null)
            textToDisplayObj = textToDisplay;

        var excelRange = anchor as ExcelRange;
        var hyperlink = _hyperlinks.Add(excelRange?.InternalRange, address, subAddressObj, screenTipObj, textToDisplayObj) as MsExcel.Hyperlink;
        return hyperlink != null ? new ExcelHyperlink(hyperlink) : null;
    }

    /// <summary>
    /// 删除集合中的所有超链接
    /// </summary>
    public void Delete()
    {
        _hyperlinks?.Delete();
    }

    /// <summary>
    /// 删除指定索引的超链接
    /// </summary>
    /// <param name="index">要删除的超链接索引</param>
    public void Delete(int index)
    {
        if (_hyperlinks != null && index >= 1 && index <= Count)
        {
            var hyperlink = _hyperlinks[index] as MsExcel.Hyperlink;
            hyperlink?.Delete();
        }
    }

    public IEnumerator<IExcelHyperlink> GetEnumerator()
    {
        if (_hyperlinks == null)
            yield break;
        for (int i = 0; i < _hyperlinks.Count; i++)
        {
            yield return new ExcelHyperlink(_hyperlinks[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
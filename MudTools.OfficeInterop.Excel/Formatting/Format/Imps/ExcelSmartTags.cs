//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// Excel SmartTags 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.SmartTags 对象的安全访问和资源管理
/// </summary>
internal class ExcelSmartTags : IExcelSmartTags
{
    /// <summary>
    /// 底层的 COM SmartTags 集合对象
    /// </summary>
    private MsExcel.SmartTags _smartTags;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelSmartTags 实例
    /// </summary>
    /// <param name="smartTags">底层的 COM SmartTags 集合对象</param>
    internal ExcelSmartTags(MsExcel.SmartTags smartTags)
    {
        _smartTags = smartTags;
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
                // 释放所有子智能标记对象
                for (int i = 1; i <= Count; i++)
                {
                    var smartTag = this[i] as ExcelSmartTag;
                    smartTag?.Dispose();
                }

                // 释放底层COM对象
                if (_smartTags != null)
                    Marshal.ReleaseComObject(_smartTags);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _smartTags = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取智能标记集合中的智能标记数量
    /// </summary>
    public int Count => _smartTags?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的智能标记对象
    /// </summary>
    /// <param name="index">智能标记索引（从1开始）</param>
    /// <returns>智能标记对象</returns>
    public IExcelSmartTag this[int index]
    {
        get
        {
            if (_smartTags == null || index < 1 || index > Count)
                return null;

            var smartTag = _smartTags[index] as MsExcel.SmartTag;
            return smartTag != null ? new ExcelSmartTag(smartTag) : null;
        }
    }

    /// <summary>
    /// 向集合中添加新的智能标记
    /// </summary>
    /// <param name="smartTagType">智能标记类型</param>
    /// <param name="text">智能标记文本</param>
    /// <returns>新创建的智能标记对象</returns>
    public IExcelSmartTag Add(string smartTagType, string text)
    {
        if (_smartTags == null || string.IsNullOrEmpty(smartTagType) || string.IsNullOrEmpty(text))
            return null;

        var smartTag = _smartTags.Add(smartTagType) as MsExcel.SmartTag;
        return smartTag != null ? new ExcelSmartTag(smartTag) : null;
    }



    public IEnumerator<IExcelSmartTag> GetEnumerator()
    {
        for (var i = 0; i < _smartTags.Count; i++)
        {
            var smartTag = _smartTags[i];
            yield return new ExcelSmartTag(smartTag);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
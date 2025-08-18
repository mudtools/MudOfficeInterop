//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// 对 Microsoft.Office.Interop.Excel.GroupShapes 的封装实现类
/// </summary>
internal class ExcelGroupShapes : IExcelGroupShapes
{
    #region 属性封装

    /// <summary>
    /// 获取组合形状的数量
    /// </summary>
    public int Count => Convert.ToInt32(_groupShapes.Count);

    /// <summary>
    /// 获取指定索引的形状对象（伪代码）
    /// </summary>
    /// <param name="index">形状索引（从1开始）或名称</param>
    /// <returns>形状对象</returns>
    public IExcelShape this[object index] => new ExcelShape(_groupShapes.Item(index));

    /// <summary>
    /// 获取父对象（伪代码）
    /// </summary>
    public object Parent => _groupShapes.Parent;


    /// <summary>
    /// 获取边框所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            var application = _groupShapes?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    /// <summary>
    /// 获取创建者信息
    /// </summary>
    public int Creator => Convert.ToInt32(_groupShapes.Creator);

    #endregion

    #region 构造函数与私有字段

    private MsExcel.GroupShapes _groupShapes;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelGroupShapes 实例
    /// </summary>
    /// <param name="groupShapes">原始 COM GroupShapes 对象</param>
    internal ExcelGroupShapes(MsExcel.GroupShapes groupShapes)
    {
        _groupShapes = groupShapes ?? throw new ArgumentNullException(nameof(groupShapes));
        _disposedValue = false;
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

        if (disposing && _groupShapes != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_groupShapes) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _groupShapes = null;
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

    public IEnumerator<IExcelShape> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return new ExcelShape(_groupShapes.Item(i));
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
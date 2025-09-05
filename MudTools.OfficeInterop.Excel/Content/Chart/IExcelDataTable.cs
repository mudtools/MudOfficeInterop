//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel DataTable 对象的二次封装接口 (图表中的数据表)
/// 提供对 Microsoft.Office.Interop.Excel.DataTable 的安全访问和操作
/// </summary>
public interface IExcelDataTable : IDisposable
{
    #region 基础属性    

    /// <summary>
    /// 获取数据表对象的父对象 (通常是 Chart)
    /// 对应 DataTable.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据表对象所在的 Application 对象
    /// 对应 DataTable.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取数据表的字体对象
    /// 对应 DataTable.Font 属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取或设置是否自动缩放字体
    /// 对应 DataTable.AutoScaleFont 属性
    /// </summary>
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取数据表的边框对象
    /// 对应 DataTable.Border 属性
    /// </summary>
    IExcelBorder Border { get; }

    /// <summary>
    /// 获取数据表的格式对象
    /// 对应 DataTable.Format 属性
    /// </summary>
    IExcelChartFormat Format { get; }

    /// <summary>
    /// 获取或设置是否在数据表中显示图例项标示
    /// 对应 DataTable.ShowLegendKey 属性
    /// </summary>
    bool ShowLegendKey { get; set; }

    /// <summary>
    /// 获取或设置是否在数据表中显示水平单元格边框
    /// 对应 DataTable.HasBorderHorizontal 属性
    /// </summary>
    bool HasBorderHorizontal { get; set; }

    /// <summary>
    /// 获取或设置是否在数据表中显示垂直单元格边框
    /// 对应 DataTable.HasBorderVertical 属性
    /// </summary>
    bool HasBorderVertical { get; set; }

    /// <summary>
    /// 获取或设置是否在数据表中显示轮廓边框
    /// 对应 DataTable.HasBorderOutline 属性
    /// </summary>
    bool HasBorderOutline { get; set; }
    #endregion


    #region 操作方法
    /// <summary>
    /// 选择数据表对象
    /// 对应 DataTable.Select 方法
    /// </summary>
    void Select();

    /// <summary>
    /// 删除数据表 (通常意味着隐藏数据表，即设置 Chart.HasDataTable = false)
    /// 对应 DataTable.Delete 方法 (如果存在，但通常不直接调用)
    /// </summary>
    void Delete();
    #endregion   

    #region 格式设置方法  

    /// <summary>
    /// 设置数据表样式
    /// </summary>
    /// <param name="hasHorizontal">是否有水平边框</param>
    /// <param name="hasVertical">是否有垂直边框</param>
    /// <param name="hasOutline">是否有轮廓边框</param>
    void SetDataTableStyle(bool hasHorizontal, bool hasVertical, bool hasOutline);
    #endregion   
}

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Gridlines 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Gridlines 的安全访问和操作
/// </summary>
public interface IExcelGridlines : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取网格线对象的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取网格线对象的父对象 (通常是 Axis)
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取网格线对象所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取网格线的边框对象
    /// </summary>
    IExcelBorder Border { get; }

    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat Format { get; }
    #endregion   

    #region 操作方法
    /// <summary>
    /// 选择网格线对象
    /// </summary>
    void Select();

    /// <summary>
    /// 删除网格线 (通常意味着隐藏网格线)
    /// </summary>
    void Delete();
    #endregion
}


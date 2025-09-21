//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel DataBarBorder 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.DataBarBorder 的安全访问和操作
/// </summary>
public interface IExcelDataBarBorder : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据条边框对象的父对象 (通常是 DataBar)
    /// 对应 DataBarBorder.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据条边框对象所在的Application对象
    /// 对应 DataBarBorder.Application 属性
    /// </summary>
    IExcelApplication Application { get; }


    /// <summary>
    /// 获取或设置数据条边框的颜色
    /// 对应 DataBarBorder.Color 属性 或 .ColorIndex
    /// </summary>
    int Color { get; }
    #endregion
}

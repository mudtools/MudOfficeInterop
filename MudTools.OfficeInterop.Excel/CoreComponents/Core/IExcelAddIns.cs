//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel AddIns 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.AddIns 的安全访问和操作
/// </summary>
public interface IExcelAddIns : IDisposable, IEnumerable<IExcelAddIn>
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取加载项集合中的加载项数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    object Application { get; }

    /// <summary>
    /// 通过索引获取加载项
    /// </summary>
    /// <param name="index">加载项索引</param>
    /// <returns>加载项对象</returns>
    IExcelAddIn this[object index] { get; }

    /// <summary>
    /// 向集合中添加新加载项
    /// </summary>
    /// <param name="filename">加载项文件名</param>
    /// <param name="copyFile">是否复制文件</param>
    /// <returns>新创建的加载项对象</returns>
    IExcelAddIn Add(string filename, object copyFile);
}
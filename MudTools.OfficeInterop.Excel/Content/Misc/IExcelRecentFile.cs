//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel RecentFile 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.RecentFile 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelRecentFile : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// 对应 RecentFile.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取最近使用文件的名称 (通常包含路径)
    /// 对应 RecentFile.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取最近使用文件在集合中的索引
    /// 对应 RecentFile.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取最近使用文件的完整路径
    /// 对应 RecentFile.Path 属性 (如果存在，或从 Name 解析)
    /// </summary>
    string Path { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 打开此最近使用的文件
    /// 对应 RecentFile.Open 方法
    /// </summary>
    /// <returns>打开的工作簿对象</returns>
    IExcelWorkbook Open();

    /// <summary>
    /// 删除此最近使用文件记录 (从 RecentFiles 集合中移除)
    /// 对应 RecentFile.Delete 方法
    /// </summary>
    void Delete();

    #endregion

}

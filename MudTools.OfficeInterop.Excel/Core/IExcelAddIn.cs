//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel AddIn 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.AddIn 的安全访问和操作
/// </summary>
public interface IExcelAddIn : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置加载项的名称
    /// 对应 AddIn.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取加载项的完整路径
    /// 对应 AddIn.FullName 属性
    /// </summary>
    string FullName { get; }

    string Title { get; }
    string Subject { get; }
    string Path { get; }
    string Comments { get; }

    string Author { get; }

    string Keywords { get; }

    /// <summary>
    /// 获取加载项的父对象（通常是 AddIns 集合）
    /// 对应 AddIn.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取加载项所在的Application对象
    /// 对应 AddIn.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取加载项的程序标识符 (ProgID)
    /// 对应 AddIn.ProgId 属性
    /// </summary>
    string ProgId { get; }

    /// <summary>
    /// 获取加载项的 CLSID
    /// 对应 AddIn.CLSID 属性
    /// </summary>
    string CLSID { get; }
    #endregion

    #region 状态属性
    /// <summary>
    /// 获取或设置加载项是否已安装
    /// 对应 AddIn.Installed 属性
    /// </summary>
    bool Installed { get; set; }

    /// <summary>
    /// 获取加载项是否已打开/加载
    /// 对应 AddIn.IsOpen 属性
    /// </summary>
    bool IsOpen { get; }

    #endregion
}

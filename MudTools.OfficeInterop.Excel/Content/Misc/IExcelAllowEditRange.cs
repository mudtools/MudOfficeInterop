//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示受保护工作表中可以编辑的单元格区域。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelAllowEditRange : IOfficeObject<IExcelAllowEditRange, MsExcel.AllowEditRange>, IDisposable
{

    /// <summary>
    /// 获取或设置将文档保存为网页时的网页标题。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置 Range 对象，该对象表示受保护工作表中可以编辑的区域子集。
    /// </summary>
    IExcelRange? Range { get; set; }

    /// <summary>
    /// 更改受保护工作表中可编辑区域的密码。
    /// </summary>
    /// <param name="password">必需。新密码。</param>
    void ChangePassword(string password);

    /// <summary>
    /// 删除该对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 取消对工作表或工作簿的保护。
    /// </summary>
    /// <param name="password">可选项。区分大小写的字符串，用于取消保护工作表或工作簿的密码。如果工作表或工作簿未使用密码保护，则忽略此参数。</param>
    void Unprotect(string? password = null);

    /// <summary>
    /// 获取工作表中受保护区域的 UserAccessList 对象。
    /// </summary>
    IExcelUserAccessList? Users { get; }
}
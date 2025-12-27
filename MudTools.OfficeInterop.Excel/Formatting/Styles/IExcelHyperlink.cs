//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Hyperlink 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Hyperlink 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelHyperlink : IOfficeObject<IExcelHyperlink>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取超链接的名称
    /// 对应 Hyperlink.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置超链接的目标地址
    /// 对应 Hyperlink.Address 属性
    /// </summary>
    string Address { get; set; }

    /// <summary>
    /// 获取或设置超链接的子地址（如工作表名称）
    /// 对应 Hyperlink.SubAddress 属性
    /// </summary>
    string SubAddress { get; set; }

    /// <summary>
    /// 获取或设置鼠标悬停时显示的提示文本
    /// 对应 Hyperlink.ScreenTip 属性
    /// </summary>
    string ScreenTip { get; set; }

    /// <summary>
    /// 获取或设置要显示的文本
    /// 对应 Hyperlink.TextToDisplay 属性
    /// </summary>
    string TextToDisplay { get; set; }

    /// <summary>
    /// 获取或设置电子邮件超链接的主题行
    /// 对应 Hyperlink.EmailSubject 属性
    /// </summary>
    string EmailSubject { get; set; }

    /// <summary>
    /// 获取与超链接关联的形状对象
    /// 对应 Hyperlink.Shape 属性
    /// </summary>
    IExcelShape? Shape { get; }

    /// <summary>
    /// 获取超链接所在的区域对象
    /// 对应 Hyperlink.Range 属性
    /// </summary>
    IExcelRange? Range { get; }

    /// <summary>
    /// 获取超链接的类型
    /// 对应 Hyperlink.Type 属性
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 删除超链接
    /// 对应 Hyperlink.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 跟随超链接（打开链接）
    /// 对应 Hyperlink.Follow 方法
    /// </summary>
    /// <param name="newWindow">是否在新窗口中打开</param>
    /// <param name="addHistory">是否添加到历史记录</param>
    /// <param name="extraInfo">额外信息</param>
    void Follow(bool newWindow, bool addHistory, object extraInfo);

    /// <summary>
    /// 创建新的文档并将其作为超链接目标
    /// </summary>
    /// <param name="filename">要创建的文件名</param>
    /// <param name="editNow">是否立即编辑新文档</param>
    /// <param name="overwrite">是否覆盖同名现有文件</param>
    void CreateNewDocument(string filename, bool editNow, bool overwrite);

    /// <summary>
    /// 将超链接添加到收藏夹
    /// </summary>
    void AddToFavorites();
}
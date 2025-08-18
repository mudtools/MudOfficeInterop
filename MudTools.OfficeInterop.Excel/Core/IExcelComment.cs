//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Comment 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Comment 的安全访问和操作
/// </summary>
public interface IExcelComment : IDisposable
{
    /// <summary>
    /// 获取或设置注释的文本内容
    /// </summary>
    string Text(string? text = null, int? start = null, bool? overwrite = null);

    /// <summary>
    /// 获取注释的作者
    /// 对应 Comment.Author 属性
    /// </summary>
    string Author { get; }

    /// <summary>
    /// 获取或设置注释是否可见
    /// 对应 Comment.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取注释所在的区域对象
    /// 对应 Comment.Parent 属性
    /// </summary>
    IExcelRange Parent { get; }

    /// <summary>
    /// 获取注释的形状对象
    /// 对应 Comment.Shape 属性
    /// </summary>
    IExcelShape Shape { get; }

    /// <summary>
    /// 删除注释
    /// 对应 Comment.Delete 方法
    /// </summary>
    void Delete();
}
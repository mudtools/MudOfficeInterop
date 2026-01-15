//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的编辑框控件接口
/// 继承自IExcelControl接口和IDisposable接口，提供编辑框特有的功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelEditBox : IOfficeObject<IExcelEditBox, MsExcel.EditBox>, IExcelControl, IDisposable
{
    /// <summary>
    /// 获取或设置编辑框的标题
    /// </summary>
    string Caption { get; set; }


    /// <summary>
    /// 获取或设置下拉框的文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 选择编辑框
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);


    /// <summary>
    /// 删除编辑框
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制编辑框
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切编辑框
    /// </summary>
    void Cut();
}
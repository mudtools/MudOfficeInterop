//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的列表框控件接口
/// 继承自IExcelControl接口和IDisposable接口，提供列表框特有的功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelListBox : IOfficeObject<IExcelListBox>, IExcelControl, IDisposable
{

    /// <summary>
    /// 获取或设置列表框的链接单元格
    /// </summary>
    string LinkedCell { get; set; }

    /// <summary>
    /// 获取或设置列表框的列表范围
    /// </summary>
    string ListFillRange { get; set; }

    /// <summary>
    /// 获取或设置下拉框的值（选中项的索引）
    /// </summary>
    int Value { get; set; }

    /// <summary>
    /// 获取列表框中的项目数量
    /// </summary>
    int ListCount { get; }

    /// <summary>
    /// 选择列表框
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除列表框
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制列表框
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切列表框
    /// </summary>
    void Cut();
}
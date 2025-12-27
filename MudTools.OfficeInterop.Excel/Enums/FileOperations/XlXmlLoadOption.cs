//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定 Excel 如何打开 XML 数据文件
/// </summary>
public enum XlXmlLoadOption
{
    /// <summary>
    /// 提示用户选择打开文件的方式
    /// </summary>
    xlXmlLoadPromptUser,

    /// <summary>
    /// 打开 XML 数据文件。文件内容将被扁平化处理
    /// </summary>
    xlXmlLoadOpenXml,

    /// <summary>
    /// 将 XML 数据文件的内容放入 XML 列表中
    /// </summary>
    xlXmlLoadImportToList,

    /// <summary>
    /// 在 XML 结构任务窗格中显示 XML 数据文件的架构
    /// </summary>
    xlXmlLoadMapXml
}
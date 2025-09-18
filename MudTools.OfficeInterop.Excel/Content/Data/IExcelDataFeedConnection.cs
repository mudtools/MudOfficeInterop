//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel数据源连接接口，用于定义数据连接的基本操作和属性
/// </summary>
public interface IExcelDataFeedConnection : IDisposable
{
    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取连接的命令文本
    /// </summary>
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    IExcelWorkbookConnection Parent { get; }

    /// <summary>
    /// 获取或设置刷新时是否提示文件未找到
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取或设置是否保存密码
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 刷新数据源连接
    /// </summary>
    void Refresh();
}

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
using System;

public interface IExcelOLEDBConnection : IDisposable
{

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置连接的命令文本
    /// </summary>
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取或设置连接字符串
    /// </summary>
    string Connection { get; set; }

    /// <summary>
    /// 获取ADO连接对象
    /// </summary>
    object ADOConnection { get; }

    /// <summary>
    /// 获取或设置连接是否启用背景刷新
    /// </summary>
    bool BackgroundQuery { get; set; }

    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    IExcelWorkbookConnection Parent { get; }

    /// <summary>
    /// 获取或设置是否启用连接
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 获取或设置刷新时是否提示用户
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取或设置是否保存密码
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 获取或设置源数据文件
    /// </summary>
    string SourceDataFile { get; set; }

    /// <summary>
    /// 获取源工作簿连接名称
    /// </summary>
    string SourceConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置是否始终使用连接文件
    /// </summary>
    bool AlwaysUseConnectionFile { get; set; }

    /// <summary>
    /// 刷新OLEDB连接
    /// </summary>
    void Refresh();

    /// <summary>
    /// 测试OLEDB连接
    /// </summary>
    /// <returns>连接是否成功</returns>
    bool TestConnection();

    /// <summary>
    /// 执行SQL命令
    /// </summary>
    /// <param name="sql">SQL命令</param>
    /// <returns>受影响的行数</returns>
    int ExecuteCommand(string sql);

}

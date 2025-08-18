//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel 窗口状态枚举
/// 用于指定Excel应用程序窗口的状态
/// </summary>
public enum XlWindowState
{
    /// <summary>
    /// 最大化窗口
    /// 窗口占据整个屏幕
    /// </summary>
    xlMaximized = -4137,
    
    /// <summary>
    /// 最小化窗口
    /// 窗口缩小为任务栏上的图标
    /// </summary>
    xlMinimized = -4140,
    
    /// <summary>
    /// 常规窗口
    /// 窗口以用户定义的大小显示
    /// </summary>
    xlNormal = -4143
}
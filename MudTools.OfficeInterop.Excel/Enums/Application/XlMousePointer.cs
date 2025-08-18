//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 鼠标指针枚举
/// 用于指定在特定操作期间显示的鼠标指针类型
/// </summary>
public enum XlMousePointer
{
    /// <summary>
    /// I型光标
    /// 文本输入光标，通常在可编辑区域显示
    /// </summary>
    xlIBeam = 3,
    
    /// <summary>
    /// 默认光标
    /// 系统默认的鼠标指针
    /// </summary>
    xlDefault = -4143,
    
    /// <summary>
    /// 西北箭头光标
    /// 指向西北方向的箭头光标
    /// </summary>
    xlNorthwestArrow = 1,
    
    /// <summary>
    /// 等待光标
    /// 沙漏或旋转圈等表示等待状态的光标
    /// </summary>
    xlWait = 2
}
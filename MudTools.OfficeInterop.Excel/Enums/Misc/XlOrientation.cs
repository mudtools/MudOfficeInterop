//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 方向枚举
/// 用于指定对象（如文字、图表元素等）的方向或排列方式
/// </summary>
public enum XlOrientation
{
    /// <summary>
    /// 向下
    /// 对象方向从上到下倾斜（如文字沿斜线向下排列）
    /// </summary>
    xlDownward = -4170,
    
    /// <summary>
    /// 水平
    /// 对象水平排列
    /// </summary>
    xlHorizontal = -4128,
    
    /// <summary>
    /// 向上
    /// 对象方向从下到上倾斜（如文字沿斜线向上排列）
    /// </summary>
    xlUpward = -4171,
    
    /// <summary>
    /// 垂直
    /// 对象垂直排列
    /// </summary>
    xlVertical = -4166
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 垂直对齐方式枚举
/// 用于指定单元格中文本的垂直对齐方式
/// </summary>
public enum VerticalAlignment
{
    /// <summary>
    /// 顶部对齐
    /// 文本与单元格顶部对齐
    /// </summary>
    Top = -4160,
    
    /// <summary>
    /// 居中对齐
    /// 文本在单元格中垂直居中对齐
    /// </summary>
    Center = -4108,
    
    /// <summary>
    /// 底部对齐
    /// 文本与单元格底部对齐
    /// </summary>
    Bottom = -4107,
    
    /// <summary>
    /// 两端对齐
    /// 文本在单元格中均匀分布，调整行间距使文本充满整个单元格高度
    /// </summary>
    Justify = -4130,
    
    /// <summary>
    /// 分散对齐
    /// 文本在单元格中均匀分布，最后一行可能左对齐、居中或右对齐
    /// </summary>
    Distributed = -4117
}
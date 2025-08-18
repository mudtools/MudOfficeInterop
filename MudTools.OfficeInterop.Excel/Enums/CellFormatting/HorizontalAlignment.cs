//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 水平对齐方式枚举
/// 用于指定单元格中文本的水平对齐方式
/// </summary>
public enum HorizontalAlignment
{
    /// <summary>
    /// 常规对齐 - 根据数据类型自动选择对齐方式，文本左对齐，数字右对齐
    /// </summary>
    General = 1,
    
    /// <summary>
    /// 左对齐 - 文本或数字靠左对齐
    /// </summary>
    Left = -4131,
    
    /// <summary>
    /// 居中对齐 - 文本或数字在单元格中央对齐
    /// </summary>
    Center = -4108,
    
    /// <summary>
    /// 右对齐 - 文本或数字靠右对齐
    /// </summary>
    Right = -4152,
    
    /// <summary>
    /// 填充对齐 - 重复文本以填充整个单元格宽度
    /// </summary>
    Fill = 5,
    
    /// <summary>
    /// 两端对齐 - 调整文本间距使文本均匀分布在单元格中
    /// </summary>
    Justify = -4130,
    
    /// <summary>
    /// 跨选定区域居中 - 在选定的多个单元格中居中对齐
    /// </summary>
    CenterAcrossSelection = 7,
    
    /// <summary>
    /// 分散对齐 - 类似两端对齐，但会调整字符间距使其均匀分布
    /// </summary>
    Distributed = -4117
}
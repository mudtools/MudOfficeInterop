//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Border 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Border 的安全访问和操作
/// </summary>
public interface IExcelBorder : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置边框的线条样式
    /// 对应 Border.LineStyle 属性
    /// </summary>
    XlLineStyle LineStyle { get; set; }

    /// <summary>
    /// 获取或设置边框的粗细
    /// 对应 Border.Weight 属性
    /// </summary>
    int Weight { get; set; }

    /// <summary>
    /// 获取或设置边框的颜色
    /// 对应 Border.Color 属性
    /// </summary>
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置边框的主题颜色
    /// 对应 Border.ThemeColor 属性
    /// </summary>
    Color ThemeColor { get; set; }

    /// <summary>
    /// 获取或设置边框的着色和阴影
    /// 对应 Border.TintAndShade 属性
    /// </summary>
    double TintAndShade { get; set; }

    /// <summary>
    /// 获取边框所在的父对象
    /// 对应 Border.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取边框所在的Application对象
    /// 对应 Border.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }
    #endregion

    #region 格式设置  

    /// <summary>
    /// 重置边框为默认值
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制边框格式
    /// </summary>
    /// <param name="sourceBorder">源边框</param>
    void CopyFormat(IExcelBorder sourceBorder);

    /// <summary>
    /// 应用预设样式
    /// </summary>
    /// <param name="presetStyle">预设样式类型</param>
    void ApplyPresetStyle(int presetStyle);

    #endregion

    #region 导出和转换

    /// <summary>
    /// 导出边框到文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="overwrite">是否覆盖已存在文件</param>
    /// <returns>是否导出成功</returns>
    bool Export(string filename, bool overwrite = true);
    #endregion

    #region 高级功能

    /// <summary>
    /// 获取边框的详细信息
    /// </summary>
    /// <returns>边框详细信息对象</returns>
    BorderDetails GetDetails();

    /// <summary>
    /// 验证边框设置
    /// </summary>
    /// <returns>验证结果</returns>
    BorderValidationResult Validate();

    /// <summary>
    /// 比较两个边框
    /// </summary>
    /// <param name="otherBorder">要比较的边框</param>
    /// <returns>比较结果</returns>
    BorderComparisonResult Compare(IExcelBorder otherBorder);

    /// <summary>
    /// 克隆边框
    /// </summary>
    /// <returns>克隆的边框对象</returns>
    IExcelBorder Clone();

    #endregion
}


/// <summary>
/// 边框验证结果结构
/// </summary>
public class BorderValidationResult
{
    /// <summary>
    /// 是否有效
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// 错误信息
    /// </summary>
    public string ErrorMessage { get; set; }

    /// <summary>
    /// 建议的修正方案
    /// </summary>
    public string SuggestedFix { get; set; }

    /// <summary>
    /// 边框类型有效性
    /// </summary>
    public bool ValidBorderType { get; set; }

    /// <summary>
    /// 线条样式有效性
    /// </summary>
    public bool ValidLineStyle { get; set; }

    /// <summary>
    /// 颜色有效性
    /// </summary>
    public bool ValidColor { get; set; }

    /// <summary>
    /// 粗细有效性
    /// </summary>
    public bool ValidWeight { get; set; }

    /// <summary>
    /// 位置有效性
    /// </summary>
    public bool ValidPosition { get; set; }

    /// <summary>
    /// 大小有效性
    /// </summary>
    public bool ValidSize { get; set; }

    /// <summary>
    /// 是否超出边界
    /// </summary>
    public bool OutOfBounds { get; set; }

    /// <summary>
    /// 是否与其他边框重叠
    /// </summary>
    public bool Overlapping { get; set; }

    /// <summary>
    /// 是否为负尺寸
    /// </summary>
    public bool NegativeDimensions { get; set; }
}

/// <summary>
/// 边框比较结果结构
/// </summary>
public class BorderComparisonResult
{
    /// <summary>
    /// 是否相等
    /// </summary>
    public bool AreEqual { get; set; }

    /// <summary>
    /// 相似度
    /// </summary>
    public double Similarity { get; set; }

    /// <summary>
    /// 差异点数组
    /// </summary>
    public string[] Differences { get; set; }

    /// <summary>
    /// 相同点数组
    /// </summary>
    public string[] Similarities { get; set; }

    /// <summary>
    /// 第一个边框名称
    /// </summary>
    public string FirstBorderName { get; set; }

    /// <summary>
    /// 第二个边框名称
    /// </summary>
    public string SecondBorderName { get; set; }

    /// <summary>
    /// 比较时间
    /// </summary>
    public DateTime ComparisonTime { get; set; }

    /// <summary>
    /// 比较类型
    /// </summary>
    public string ComparisonType { get; set; }

    /// <summary>
    /// 比较结果描述
    /// </summary>
    public string Description { get; set; }

    /// <summary>
    /// 建议的操作
    /// </summary>
    public string RecommendedAction { get; set; }

    /// <summary>
    /// 是否可以合并
    /// </summary>
    public bool CanBeMerged { get; set; }

    /// <summary>
    /// 是否可以替换
    /// </summary>
    public bool CanBeReplaced { get; set; }

    /// <summary>
    /// 是否可以继承
    /// </summary>
    public bool CanBeInherited { get; set; }

    /// <summary>
    /// 是否可以覆盖
    /// </summary>
    public bool CanBeOverridden { get; set; }

    /// <summary>
    /// 是否需要更新
    /// </summary>
    public bool NeedsUpdate { get; set; }

    /// <summary>
    /// 是否需要同步
    /// </summary>
    public bool NeedsSynchronization { get; set; }

    /// <summary>
    /// 是否需要验证
    /// </summary>
    public bool NeedsValidation { get; set; }

    /// <summary>
    /// 是否需要优化
    /// </summary>
    public bool NeedsOptimization { get; set; }
}

/// <summary>
/// 边框详细信息结构
/// </summary>
public class BorderDetails
{
    /// <summary>
    /// 边框类型
    /// </summary>
    public int BorderType { get; set; }

    /// <summary>
    /// 边框索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 边框名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 线条样式
    /// </summary>
    public int LineStyle { get; set; }

    /// <summary>
    /// 线条样式名称
    /// </summary>
    public string LineStyleName { get; set; }

    /// <summary>
    /// 边框粗细
    /// </summary>
    public int Weight { get; set; }

    /// <summary>
    /// 边框粗细名称
    /// </summary>
    public string WeightName { get; set; }

    /// <summary>
    /// 边框颜色
    /// </summary>
    public int Color { get; set; }

    /// <summary>
    /// 边框颜色名称
    /// </summary>
    public string ColorName { get; set; }

    /// <summary>
    /// 主题颜色
    /// </summary>
    public int ThemeColor { get; set; }

    /// <summary>
    /// 着色和阴影
    /// </summary>
    public double TintAndShade { get; set; }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool Visible { get; set; }

    /// <summary>
    /// 左边距
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// 顶边距
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// 宽度
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// 旋转角度
    /// </summary>
    public double Rotation { get; set; }

    /// <summary>
    /// 是否锁定
    /// </summary>
    public bool Locked { get; set; }

    /// <summary>
    /// 是否为实线
    /// </summary>
    public bool IsSolid { get; set; }

    /// <summary>
    /// 是否为虚线
    /// </summary>
    public bool IsDash { get; set; }

    /// <summary>
    /// 是否为点线
    /// </summary>
    public bool IsDot { get; set; }

    /// <summary>
    /// 是否为双线
    /// </summary>
    public bool IsDouble { get; set; }

    /// <summary>
    /// 是否为主要颜色
    /// </summary>
    public bool IsPrimaryColor { get; set; }

    /// <summary>
    /// 是否为自定义颜色
    /// </summary>
    public bool IsCustomColor { get; set; }

    /// <summary>
    /// 是否为细线
    /// </summary>
    public bool IsThin { get; set; }

    /// <summary>
    /// 是否为中等线
    /// </summary>
    public bool IsMedium { get; set; }

    /// <summary>
    /// 是否为粗线
    /// </summary>
    public bool IsThick { get; set; }

    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime Created { get; set; }

    /// <summary>
    /// 修改时间
    /// </summary>
    public DateTime Modified { get; set; }

    /// <summary>
    /// 所属工作表
    /// </summary>
    public string Worksheet { get; set; }

    /// <summary>
    /// 所属区域
    /// </summary>
    public string RangeAddress { get; set; }

    /// <summary>
    /// 边框类别
    /// </summary>
    public string Category { get; set; }

    /// <summary>
    /// 边框描述
    /// </summary>
    public string Description { get; set; }

    /// <summary>
    /// 边框标签
    /// </summary>
    public string[] Tags { get; set; }

    /// <summary>
    /// 边框优先级
    /// </summary>
    public int Priority { get; set; }

    /// <summary>
    /// 是否启用
    /// </summary>
    public bool IsEnabled { get; set; }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool IsVisible { get; set; }

    /// <summary>
    /// 边框透明度
    /// </summary>
    public int Transparency { get; set; }

    /// <summary>
    /// 边框渐变类型
    /// </summary>
    public int GradientType { get; set; }

    /// <summary>
    /// 边框渐变角度
    /// </summary>
    public int GradientAngle { get; set; }

    /// <summary>
    /// 边框渐变颜色1
    /// </summary>
    public int GradientColor1 { get; set; }

    /// <summary>
    /// 边框渐变颜色2
    /// </summary>
    public int GradientColor2 { get; set; }

    /// <summary>
    /// 边框渐变停止点1
    /// </summary>
    public double GradientStop1 { get; set; }

    /// <summary>
    /// 边框渐变停止点2
    /// </summary>
    public double GradientStop2 { get; set; }
}

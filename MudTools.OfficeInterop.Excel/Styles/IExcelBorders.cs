//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// Excel Borders 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Borders 的安全访问和操作
/// </summary>
public interface IExcelBorders : IEnumerable<IExcelBorder>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 应用到全局。
    /// </summary>
    bool ApplyToAll { get; set; }

    /// <summary>
    /// 获取边框集合中的边框数量
    /// 对应 Borders.Count 属性
    /// </summary>
    int Count { get; }

    Dictionary<XlBordersIndex, IExcelCellFormat> CustomBorders { get; set; }

    /// <summary>
    /// 获取指定类型的边框对象
    /// </summary>
    /// <param name="borderType">边框类型</param>
    /// <returns>边框对象</returns>
    IExcelBorder this[XlBordersIndex borderType] { get; }

    /// <summary>
    /// 获取边框集合所在的父对象
    /// 对应 Borders.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取边框集合所在的Application对象
    /// 对应 Borders.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置边框线条样式
    /// </summary>
    XlLineStyle LineStyle { get; set; }

    /// <summary>
    /// 获取或设置边框粗细
    /// </summary>
    XlBorderWeight Weight { get; set; }

    /// <summary>
    /// 获取或设置字体颜色（RGB值）
    /// </summary>
    Color Color { get; set; }

    #endregion   

    #region 查找和筛选

    /// <summary>
    /// 根据线条样式查找边框
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <returns>匹配的边框数组</returns>
    IExcelBorder[] FindByLineStyle(int lineStyle);

    /// <summary>
    /// 根据颜色查找边框
    /// </summary>
    /// <param name="color">边框颜色</param>
    /// <returns>匹配的边框数组</returns>
    IExcelBorder[] FindByColor(int color);

    /// <summary>
    /// 根据粗细查找边框
    /// </summary>
    /// <param name="weight">边框粗细</param>
    /// <returns>匹配的边框数组</returns>
    IExcelBorder[] FindByWeight(int weight);
    #endregion

    #region 格式设置

    /// <summary>
    /// 设置所有边框的线条样式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <param name="weight">边框粗细</param>
    void SetLineStyle(XlLineStyle lineStyle, int weight = 1);

    /// <summary>
    /// 设置所有边框的颜色
    /// </summary>
    /// <param name="color">边框颜色</param>
    void SetColor(Color color);

    /// <summary>
    /// 设置所有边框的粗细
    /// </summary>
    /// <param name="weight">边框粗细</param>
    void SetWeight(int weight);

    /// <summary>
    /// 统一所有边框的格式
    /// </summary>
    /// <param name="lineStyle">线条样式</param>
    /// <param name="color">边框颜色</param>
    /// <param name="weight">边框粗细</param>
    void UniformFormat(Color color, XlLineStyle lineStyle = XlLineStyle.xlLineStyleNone, int weight = 2);

    /// <summary>
    /// 复制边框格式
    /// </summary>
    /// <param name="sourceBorder">源边框</param>
    /// <param name="applyToAll">是否应用到所有边框</param>
    void CopyFormat(IExcelBorder sourceBorder, bool applyToAll = false);

    /// <summary>
    /// 应用预设边框样式
    /// </summary>
    /// <param name="presetStyle">预设样式类型</param>
    void ApplyPresetStyle(int presetStyle);

    #endregion   

    #region 导出和导入

    /// <summary>
    /// 导出所有边框到文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <returns>是否导出成功</returns>
    bool ExportToFile(string filename);

    /// <summary>
    /// 从文件导入边框
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的边框数量</returns>
    int ImportFromFile(string filename);

    /// <summary>
    /// 获取所有边框的信息
    /// </summary>
    /// <returns>边框信息数组</returns>
    BorderInfo[] GetAllBorderInfo();
    #endregion

    #region 统计和分析   

    /// <summary>
    /// 获取线条样式统计
    /// </summary>
    /// <returns>线条样式统计信息数组</returns>
    LineStyleStatistics[] GetLineStyleStatistics();

    /// <summary>
    /// 获取颜色统计
    /// </summary>
    /// <returns>颜色统计信息</returns>
    BorderColorStatistics GetColorStatistics();

    /// <summary>
    /// 获取粗细统计
    /// </summary>
    /// <returns>粗细统计信息</returns>
    WeightStatistics GetWeightStatistics();

    #endregion

    #region 高级功能
    /// <summary>
    /// 重置边框为默认值
    /// </summary>
    void Reset();

    /// <summary>
    /// 验证边框设置
    /// </summary>
    /// <returns>验证结果</returns>
    BorderValidationResult Validate();
    #endregion
}


/// <summary>
/// 边框数据结构
/// </summary>
public class BorderData
{
    /// <summary>
    /// 边框类型
    /// </summary>
    public int BorderType { get; set; }

    /// <summary>
    /// 线条样式
    /// </summary>
    public int LineStyle { get; set; }

    /// <summary>
    /// 边框粗细
    /// </summary>
    public int Weight { get; set; }

    /// <summary>
    /// 边框颜色
    /// </summary>
    public int Color { get; set; }

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
    /// 是否可见
    /// </summary>
    public bool Visible { get; set; }

    /// <summary>
    /// 旋转角度
    /// </summary>
    public double Rotation { get; set; }
}


/// <summary>
/// 边框信息结构
/// </summary>
public class BorderInfo
{
    /// <summary>
    /// 边框索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 边框类型
    /// </summary>
    public int BorderType { get; set; }

    /// <summary>
    /// 边框名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 线条样式
    /// </summary>
    public int LineStyle { get; set; }

    /// <summary>
    /// 边框粗细
    /// </summary>
    public int Weight { get; set; }

    /// <summary>
    /// 边框颜色
    /// </summary>
    public int Color { get; set; }

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
    /// 创建时间
    /// </summary>
    public DateTime Created { get; set; }

    /// <summary>
    /// 修改时间
    /// </summary>
    public DateTime Modified { get; set; }

    /// <summary>
    /// 是否锁定
    /// </summary>
    public bool Locked { get; set; }

    /// <summary>
    /// 所属工作表
    /// </summary>
    public string Worksheet { get; set; }

    /// <summary>
    /// 所属区域
    /// </summary>
    public string RangeAddress { get; set; }
}


/// <summary>
/// 线条样式统计信息结构
/// </summary>
public class LineStyleStatistics
{
    /// <summary>
    /// 线条样式
    /// </summary>
    public int LineStyle { get; set; }

    /// <summary>
    /// 边框数量
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    /// 占比
    /// </summary>
    public double Percentage { get; set; }

    /// <summary>
    /// 样式名称
    /// </summary>
    public string StyleName { get; set; }

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
}


/// <summary>
/// 粗细统计信息结构
/// </summary>
public class WeightStatistics
{
    /// <summary>
    /// 边框粗细
    /// </summary>
    public int Weight { get; set; }

    /// <summary>
    /// 边框数量
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    /// 占比
    /// </summary>
    public double Percentage { get; set; }

    /// <summary>
    /// 粗细名称
    /// </summary>
    public string WeightName { get; set; }

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
    /// 是否为超粗线
    /// </summary>
    public bool IsExtraThick { get; set; }
}



/// <summary>
/// 颜色统计信息结构
/// </summary>
public class BorderColorStatistics
{
    /// <summary>
    /// 颜色值
    /// </summary>
    public int Color { get; set; }

    /// <summary>
    /// 颜色名称
    /// </summary>
    public string ColorName { get; set; }

    /// <summary>
    /// 边框数量
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    /// 占比
    /// </summary>
    public double Percentage { get; set; }

    /// <summary>
    /// 是否为主要颜色
    /// </summary>
    public bool IsPrimary { get; set; }

    /// <summary>
    /// 是否为自定义颜色
    /// </summary>
    public bool IsCustom { get; set; }

    /// <summary>
    /// 亮度值
    /// </summary>
    public double Brightness { get; set; }

    /// <summary>
    /// 饱和度值
    /// </summary>
    public double Saturation { get; set; }

    /// <summary>
    /// 色调值
    /// </summary>
    public double Hue { get; set; }
}
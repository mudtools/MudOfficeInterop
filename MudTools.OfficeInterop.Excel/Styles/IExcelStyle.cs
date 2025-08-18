//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Style 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Style 的安全访问和操作
/// </summary>
public interface IExcelStyle : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取本地化样式名称
    /// 对应 Style.NameLocal 属性
    /// </summary>
    public string NameLocal { get; }

    /// <summary>
    /// 获取样式名称
    /// 对应 Style.Name 属性
    /// </summary>
    string Name { get; }


    /// <summary>
    /// 获取样式是否为内置样式
    /// 对应 Style.BuiltIn 属性
    /// </summary>
    bool BuiltIn { get; }


    /// <summary>
    /// 获取样式所在的父对象
    /// 对应 Style.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取样式所在的Application对象
    /// 对应 Style.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }

    #endregion

    #region 格式属性

    /// <summary>
    /// 获取或设置是否包含数字格式
    /// </summary>
    bool IncludeNumber { get; set; }

    /// <summary>
    /// 获取或设置是否包含字体格式
    /// </summary>
    bool IncludeFont { get; set; }

    /// <summary>
    /// 获取或设置是否包含对齐格式
    /// </summary>
    bool IncludeAlignment { get; set; }

    /// <summary>
    /// 获取或设置是否添加缩进
    /// </summary>
    bool AddIndent { get; set; }


    /// <summary>
    /// 获取样式的字体对象
    /// 对应 Style.Font 属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取样式的边框对象
    /// 对应 Style.Borders 属性
    /// </summary>
    IExcelBorders Borders { get; }

    /// <summary>
    /// 获取样式的内部格式对象
    /// 对应 Style.Interior 属性
    /// </summary>
    IExcelInterior Interior { get; }

    /// <summary>
    /// 获取样式的本地化数字格式
    /// 对应 Style.NumberFormatLocal 属性
    /// </summary>
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取样式的数字格式
    /// 对应 Style.NumberFormat 属性
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取样式的水平对齐方式
    /// 对应 Style.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取样式的垂直对齐方式
    /// 对应 Style.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取样式是否自动换行
    /// 对应 Style.WrapText 属性
    /// </summary>
    bool WrapText { get; set; }

    /// <summary>
    /// 获取样式的缩进级别
    /// 对应 Style.IndentLevel 属性
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取样式的阅读顺序
    /// 对应 Style.ReadingOrder 属性
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取样式的旋转角度
    /// 对应 Style.Orientation 属性
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取样式是否添加前缀
    /// 对应 Style.ShrinkToFit 属性
    /// </summary>
    bool ShrinkToFit { get; set; }

    /// <summary>
    /// 获取样式是否合并单元格
    /// 对应 Style.MergeCells 属性
    /// </summary>
    bool MergeCells { get; set; }

    /// <summary>
    /// 获取样式是否锁定
    /// 对应 Style.Locked 属性
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取样式是否隐藏公式
    /// 对应 Style.FormulaHidden 属性
    /// </summary>
    bool FormulaHidden { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除样式
    /// 对应 Style.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制样式
    /// </summary>
    /// <param name="newName">新样式名称</param>
    /// <returns>复制的样式对象</returns>
    IExcelStyle Copy(string newName);

    /// <summary>
    /// 重命名样式
    /// </summary>
    /// <param name="newName">新样式名称</param>
    void Rename(string newName);

    /// <summary>
    /// 应用样式到指定区域
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <param name="includeFont">是否包含字体</param>
    /// <param name="includeBorder">是否包含边框</param>
    /// <param name="includeFill">是否包含填充</param>
    void ApplyTo(IExcelRange range, bool includeFont = true, bool includeBorder = true, bool includeFill = true);

    /// <summary>
    /// 更新样式
    /// </summary>
    void Update();

    /// <summary>
    /// 刷新样式
    /// </summary>
    void Refresh();

    /// <summary>
    /// 重置样式为默认值
    /// </summary>
    void Reset();

    #endregion

    #region 高级功能

    /// <summary>
    /// 克隆样式
    /// </summary>
    /// <returns>克隆的样式对象</returns>
    IExcelStyle Clone();

    /// <summary>
    /// 导出样式到文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <returns>是否导出成功</returns>
    bool Export(string filename);

    /// <summary>
    /// 从文件导入样式
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>是否导入成功</returns>
    bool Import(string filename);

    #endregion
}
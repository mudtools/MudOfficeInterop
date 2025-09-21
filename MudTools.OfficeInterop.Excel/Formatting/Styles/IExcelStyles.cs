//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Styles 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Styles 的安全访问和操作
/// </summary>
public interface IExcelStyles : IEnumerable<IExcelStyle>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取样式集合中的样式数量
    /// 对应 Styles.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的样式对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">样式索引（从1开始）</param>
    /// <returns>样式对象</returns>
    IExcelStyle? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的样式对象
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>样式对象</returns>
    IExcelStyle? this[string name] { get; }

    /// <summary>
    /// 获取样式集合所在的父对象（通常是工作簿）
    /// 对应 Styles.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取样式集合所在的Application对象
    /// 对应 Styles.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 添加新的样式
    /// 对应 Styles.Add 方法
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>新创建的样式对象</returns>
    IExcelStyle? Add(string name);

    /// <summary>
    /// 基于现有样式创建新样式
    /// </summary>
    /// <param name="name">新样式名称</param>
    /// <param name="basedOn">基础样式</param>
    /// <returns>新创建的样式对象</returns>
    IExcelStyle? AddBasedOn(string name, IExcelStyle basedOn);

    /// <summary>
    /// 批量添加样式
    /// </summary>
    /// <param name="styleNames">样式名称数组</param>
    /// <returns>成功添加的样式数量</returns>
    int AddRange(string[] styleNames);

    /// <summary>
    /// 将另一个工作簿中的样式合并到集合中
    /// </summary>
    /// <param name="workbook">目标工作簿对象</param>
    void Merge(IExcelWorkbook workbook);
    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找样式
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的样式数组</returns>
    IExcelStyle[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据字体查找样式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否粗体</param>
    /// <param name="italic">是否斜体</param>
    /// <returns>匹配的样式数组</returns>
    IExcelStyle[] FindByFont(string fontName = "", double fontSize = 0,
                            bool bold = false, bool italic = false);

    /// <summary>
    /// 根据颜色查找样式
    /// </summary>
    /// <param name="foregroundColor">前景色</param>
    /// <param name="backgroundColor">背景色</param>
    /// <param name="pattern">图案类型</param>
    /// <returns>匹配的样式数组</returns>
    IExcelStyle[] FindByColor(Color? foregroundColor = null, Color? backgroundColor = null, int pattern = -1);

    /// <summary>
    /// 根据边框查找样式
    /// </summary>
    /// <param name="borderStyle">边框样式</param>
    /// <param name="borderColor">边框颜色</param>
    /// <param name="borderWeight">边框粗细</param>
    /// <returns>匹配的样式数组</returns>
    IExcelStyle[] FindByBorder(int borderStyle = -1, int borderColor = -1, int borderWeight = -1);

    /// <summary>
    /// 获取内置样式
    /// </summary>
    /// <returns>内置样式数组</returns>
    IExcelStyle[] GetBuiltInStyles();

    /// <summary>
    /// 获取自定义样式
    /// </summary>
    /// <returns>自定义样式数组</returns>
    IExcelStyle[] GetCustomStyles();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有自定义样式
    /// 对应 Styles.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的样式
    /// </summary>
    /// <param name="index">要删除的样式索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的样式
    /// </summary>
    /// <param name="name">要删除的样式名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的样式对象
    /// </summary>
    /// <param name="style">要删除的样式对象</param>
    void Delete(IExcelStyle style);

    /// <summary>
    /// 批量删除样式
    /// </summary>
    /// <param name="names">要删除的样式名称数组</param>
    void DeleteRange(string[] names);

    /// <summary>
    /// 重命名样式
    /// </summary>
    /// <param name="oldName">旧样式名称</param>
    /// <param name="newName">新样式名称</param>
    /// <returns>是否重命名成功</returns>
    bool Rename(string oldName, string newName);

    /// <summary>
    /// 复制样式
    /// </summary>
    /// <param name="sourceStyle">源样式</param>
    /// <param name="targetName">目标样式名称</param>
    /// <returns>复制的样式对象</returns>
    IExcelStyle Copy(IExcelStyle sourceStyle, string targetName);


    #endregion

}


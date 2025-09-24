//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Workbooks 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Workbooks 的安全访问和操作
/// </summary>
public interface IExcelWorkbooks : IEnumerable<IExcelWorkbook>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取工作簿集合中的工作簿数量
    /// 对应 Workbooks.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的工作簿对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">工作簿索引（从1开始）</param>
    /// <returns>工作簿对象</returns>
    IExcelWorkbook? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的工作簿对象
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <returns>工作簿对象</returns>
    IExcelWorkbook? this[string name] { get; }

    /// <summary>
    /// 获取工作簿集合所在的父对象（通常是Application）
    /// 对应 Workbooks.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取工作簿集合所在的Application对象
    /// 对应 Workbooks.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }

    #endregion

    #region 创建和打开

    /// <summary>
    /// 打开工作簿
    /// 对应 Workbooks.Open 方法
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="updateLinks">是否更新链接</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="format">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="ignoreReadOnlyRecommended">是否忽略只读建议</param>
    /// <param name="origin">文本来源</param>
    /// <param name="delimiter">文本分隔符</param>
    /// <param name="editable">是否可编辑</param>
    /// <param name="notify">是否通知</param>
    /// <param name="converter">格式转换器</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <param name="local">是否本地</param>
    /// <param name="corruptLoad">是否损坏加载</param>
    /// <returns>打开的工作簿对象</returns>
    IExcelWorkbook? Open(string filename, int? updateLinks = 0, bool? readOnly = false,
                       int? format = 1, string? password = null, string? writeResPassword = null,
                       bool? ignoreReadOnlyRecommended = false, XlPlatform? origin = null,
                       string? delimiter = ",", bool? editable = null, bool? notify = null,
                       int? converter = null, bool? addToMru = false, bool? local = null,
                        XlCorruptLoad? corruptLoad = XlCorruptLoad.xlNormalLoad);

    /// <summary>
    /// 新建工作簿
    /// 对应 Workbooks.Add 方法
    /// </summary>
    /// <param name="template">模板文件路径</param>
    /// <returns>新建的工作簿对象</returns>
    IExcelWorkbook? Add(string template = "");


    /// <summary>
    /// 新建工作簿
    /// 对应 Workbooks.Add 方法
    /// </summary>
    /// <param name="template">模板类型</param>
    /// <returns>新建的工作簿对象</returns>
    IExcelWorkbook? Add(XlWBATemplate template);

    /// <summary>
    /// 打开工作簿（简化版本）
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="password">密码</param>
    /// <returns>打开的工作簿对象</returns>
    IExcelWorkbook? OpenSimple(string filename, bool readOnly = false, string password = "");

    /// <summary>
    /// 批量打开工作簿
    /// </summary>
    /// <param name="filenames">文件路径数组</param>
    /// <param name="readOnly">是否只读</param>
    /// <returns>成功打开的工作簿数量</returns>
    int OpenRange(string[] filenames, bool readOnly = false);

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找工作簿
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的工作簿数组</returns>
    IExcelWorkbook[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据路径查找工作簿
    /// </summary>
    /// <param name="path">文件路径</param>
    /// <returns>匹配的工作簿数组</returns>
    IExcelWorkbook[] FindByPath(string path);


    /// <summary>
    /// 获取已保存的工作簿
    /// </summary>
    /// <returns>已保存工作簿数组</returns>
    IExcelWorkbook[] GetSavedWorkbooks();

    /// <summary>
    /// 获取未保存的工作簿
    /// </summary>
    /// <returns>未保存工作簿数组</returns>
    IExcelWorkbook[] GetUnsavedWorkbooks();

    /// <summary>
    /// 获取受保护的工作簿
    /// </summary>
    /// <returns>受保护工作簿数组</returns>
    IExcelWorkbook[] GetProtectedWorkbooks();

    /// <summary>
    /// 获取只读工作簿
    /// </summary>
    /// <returns>只读工作簿数组</returns>
    IExcelWorkbook[] GetReadOnlyWorkbooks();

    #endregion

    #region 操作方法

    /// <summary>
    /// 关闭所有工作簿
    /// 对应 Workbooks.Close 方法
    /// </summary>
    void CloseAll();

    /// <summary>
    /// 删除指定索引的工作簿
    /// </summary>
    /// <param name="index">要删除的工作簿索引</param>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(int index, bool saveChanges = true);

    /// <summary>
    /// 删除指定名称的工作簿
    /// </summary>
    /// <param name="name">要删除的工作簿名称</param>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(string name, bool saveChanges = true);

    /// <summary>
    /// 删除指定的工作簿
    /// </summary>
    /// <param name="workbook">要删除的工作簿对象</param>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(IExcelWorkbook workbook, bool saveChanges = true);

    /// <summary>
    /// 批量关闭工作簿
    /// </summary>
    /// <param name="names">要关闭的工作簿名称数组</param>
    /// <param name="saveChanges">是否保存更改</param>
    void CloseRange(string[] names, bool saveChanges = true);

    /// <summary>
    /// 保存所有工作簿
    /// </summary>
    void SaveAll();

    /// <summary>
    /// 保存指定工作簿
    /// </summary>
    /// <param name="workbook">要保存的工作簿</param>
    void Save(IExcelWorkbook workbook);

    /// <summary>
    /// 另存为所有工作簿
    /// </summary>
    /// <param name="folderPath">保存文件夹路径</param>
    /// <param name="fileFormat">文件格式</param>
    /// <returns>成功保存的工作簿数量</returns>
    int SaveAsAll(string folderPath, string fileFormat = "xlsx");

    #endregion


    #region 高级功能

    /// <summary>
    /// 获取活动工作簿
    /// </summary>
    /// <returns>活动工作簿对象</returns>
    IExcelWorkbook? ActiveWorkbook { get; }

    /// <summary>
    /// 获取ThisWorkbook
    /// </summary>
    /// <returns>ThisWorkbook对象</returns>
    IExcelWorkbook? ThisWorkbook { get; }

    /// <summary>
    /// 打印所有工作簿
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOutAll(bool preview = false);

    /// <summary>
    /// 计算所有工作簿
    /// </summary>
    void CalculateAll();

    /// <summary>
    /// 刷新所有工作簿
    /// </summary>
    void RefreshAll();

    /// <summary>
    /// 保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    void ProtectAll(string password = "");

    /// <summary>
    /// 取消保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    void UnprotectAll(string password = "");

    #endregion
}
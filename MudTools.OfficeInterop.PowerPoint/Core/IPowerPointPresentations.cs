//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint Presentations 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.Presentations 的安全访问和操作
/// </summary>
public interface IPowerPointPresentations : IEnumerable<IPowerPointPresentation>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取演示文稿集合中的演示文稿数量
    /// 对应 Presentations.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的演示文稿对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">演示文稿索引（从1开始）</param>
    /// <returns>演示文稿对象</returns>
    IPowerPointPresentation this[int index] { get; }


    /// <summary>
    /// 获取演示文稿集合所在的父对象（通常是 Application）
    /// 对应 Presentations.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取演示文稿集合所在的Application对象
    /// 对应 Presentations.Application 属性
    /// </summary>
    IPowerPointApplication Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的空演示文稿
    /// 对应 Presentations.Add 方法
    /// </summary>
    /// <param name="withWindow">是否在新窗口中打开</param>
    /// <returns>新创建的演示文稿对象</returns>
    IPowerPointPresentation Add(bool withWindow = true);

    /// <summary>
    /// 打开一个现有的演示文稿文件
    /// 对应 Presentations.Open 方法
    /// </summary>
    /// <param name="fileName">演示文稿文件路径</param>
    /// <param name="readOnly">是否以只读方式打开</param>
    /// <param name="untitled">是否以无标题方式打开</param>
    /// <param name="withWindow">是否在新窗口中打开</param>
    /// <returns>打开的演示文稿对象</returns>
    IPowerPointPresentation Open(string fileName, bool readOnly = false, bool untitled = false, bool withWindow = true);

    /// <summary>
    /// 基于模板创建新的演示文稿
    /// 对应 Presentations.Open 方法 (使用模板路径)
    /// </summary>
    /// <param name="templatePath">模板文件路径 (.potx)</param>
    /// <param name="withWindow">是否在新窗口中打开</param>
    /// <returns>新创建的演示文稿对象</returns>
    IPowerPointPresentation CreateFromTemplate(string templatePath, bool withWindow = true);

    /// <summary>
    /// 批量打开演示文稿
    /// </summary>
    /// <param name="filePaths">演示文稿文件路径数组</param>
    /// <param name="readOnly">是否以只读方式打开</param>
    /// <returns>成功打开的演示文稿数量</returns>
    int OpenRange(string[] filePaths, bool readOnly = false);
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找演示文稿
    /// </summary>
    /// <param name="name">演示文稿名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的演示文稿数组</returns>
    IPowerPointPresentation[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据路径查找演示文稿
    /// </summary>
    /// <param name="path">演示文稿完整路径</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的演示文稿数组</returns>
    IPowerPointPresentation[] FindByPath(string path, bool matchCase = false);

    /// <summary>
    /// 获取活动的演示文稿
    /// </summary>
    /// <returns>活动演示文稿对象</returns>
    IPowerPointPresentation GetActivePresentation();

    #endregion

    #region 操作方法
    /// <summary>
    /// 关闭所有演示文稿
    /// </summary>
    /// <param name="saveChanges">关闭时是否保存更改</param>
    void Clear(bool saveChanges = true);

    /// <summary>
    /// 删除指定名称的演示文稿
    /// </summary>
    /// <param name="name">要关闭的演示文稿名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的演示文稿对象
    /// </summary>
    /// <param name="presentation">要关闭的演示文稿对象</param>
    /// <param name="saveChanges">关闭时是否保存更改</param>
    void Delete(IPowerPointPresentation presentation, bool saveChanges = true);

    #endregion

    #region 导出和导入 (概念性)
    /// <summary>
    /// 导出所有演示文稿到文件夹
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="format">导出格式 (例如 "pptx", "pdf")</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的演示文稿数量</returns>
    int ExportToFolder(string folderPath, PpSaveAsFileType format = PpSaveAsFileType.ppSaveAsOpenXMLPresentation, string prefix = "presentation_");

    /// <summary>
    /// 从文件夹导入演示文稿
    /// </summary>
    /// <param name="folderPath">导入文件夹路径</param>
    /// <param name="fileExtension">要导入的文件扩展名 (例如 ".pptx")</param>
    /// <returns>成功导入的演示文稿数量</returns>
    int ImportFromFolder(string folderPath, string fileExtension = ".pptx");
    #endregion
}

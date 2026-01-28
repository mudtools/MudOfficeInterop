//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 演示文稿接口（精简版）
/// </summary>
public interface IPowerPointPresentation : IDisposable
{
    /// <summary>
    /// 获取或设置演示文稿名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取演示文稿完整路径
    /// </summary>
    string FullName { get; }

    /// <summary>
    /// 获取演示文稿路径
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取幻灯片数量
    /// </summary>
    int SlideCount { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片集合
    /// </summary>
    IPowerPointSlides Slides { get; }

    /// <summary>
    /// 获取或设置演示文稿是否已修改
    /// </summary>
    bool Saved { get; set; }

    /// <summary>
    /// 获取或设置演示文稿是否只读
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 保存演示文稿
    /// </summary>
    /// <param name="fileName">文件名（可选）</param>
    /// <param name="fileFormat">文件格式（可选）</param>
    void Save(string fileName = null, PpSaveAsFileType fileFormat = PpSaveAsFileType.ppSaveAsDefault);

    /// <summary>
    /// 另存为演示文稿
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="embedTrueTypeFonts">是否嵌入TrueType字体</param>
    void SaveAs(string fileName, PpSaveAsFileType fileFormat = PpSaveAsFileType.ppSaveAsDefault, bool embedTrueTypeFonts = false);

    /// <summary>
    /// 关闭演示文稿
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(bool saveChanges = true);

    /// <summary>
    /// 导出演示文稿
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="exportFormat">导出格式</param>
    /// <param name="scaleWidth">缩放宽度</param>
    /// <param name="scaleHeight">缩放高度</param>
    void Export(string fileName, string exportFormat = "PNG", int scaleWidth = 0, int scaleHeight = 0);

    /// <summary>
    /// 保护演示文稿
    /// </summary>
    /// <param name="password">密码</param>
    /// <param name="writePassword">写入密码</param>
    /// <param name="readOnlyRecommended">是否推荐只读打开</param>
    void Protect(string password, string writePassword = null, bool readOnlyRecommended = false);

    /// <summary>
    /// 取消保护演示文稿
    /// </summary>
    /// <param name="password">密码</param>
    void Unprotect(string password);

    /// <summary>
    /// 添加幻灯片到演示文稿
    /// </summary>
    /// <param name="layout">幻灯片布局</param>
    /// <param name="position">插入位置</param>
    /// <returns>新添加的幻灯片</returns>
    IPowerPointSlide AddSlide(PpSlideLayout layout = PpSlideLayout.ppLayoutText, int position = -1);

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    void RemoveSlide(int index);

    /// <summary>
    /// 根据索引获取幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    /// <returns>幻灯片对象</returns>
    IPowerPointSlide GetSlide(int index);

    /// <summary>
    /// 获取所有幻灯片
    /// </summary>
    /// <returns>幻灯片列表</returns>
    IEnumerable<IPowerPointSlide> GetAllSlides();

    /// <summary>
    /// 获取演示文稿信息
    /// </summary>
    /// <returns>演示文稿信息字符串</returns>
    string GetPresentationInfo();

    /// <summary>
    /// 替换演示文稿中的文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    int ReplaceText(string findText, string replaceText);
}
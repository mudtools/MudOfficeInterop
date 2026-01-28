//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint SlideRange 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.SlideRange 的安全访问和操作
/// </summary>
public interface IPowerPointSlideRange : IEnumerable<IPowerPointSlide>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取幻灯片范围中的幻灯片数量
    /// 对应 SlideRange.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的幻灯片对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">幻灯片索引（从1开始）</param>
    /// <returns>幻灯片对象</returns>
    IPowerPointSlide this[int index] { get; }

    /// <summary>
    /// 获取幻灯片范围所在的父对象（通常是 Slides 集合）
    /// 对应 SlideRange.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片范围所在的Application对象
    /// 对应 SlideRange.Application 属性
    /// </summary>
    IPowerPointApplication Application { get; }

    /// <summary>
    /// 获取幻灯片范围的名称
    /// 对应 SlideRange.Name 属性
    /// </summary>
    string Name { get; set; }
    #endregion


    #region 操作方法
    /// <summary>
    /// 选择幻灯片范围中的所有幻灯片
    /// 对应 SlideRange.Select 方法
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制幻灯片范围
    /// 对应 SlideRange.Copy 方法
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切幻灯片范围
    /// 对应 SlideRange.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除幻灯片范围中的所有幻灯片
    /// 对应 SlideRange.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 删除指定索引的幻灯片
    /// </summary>
    /// <param name="index">要删除的幻灯片索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的幻灯片对象
    /// </summary>
    /// <param name="slide">要删除的幻灯片对象</param>
    void Delete(IPowerPointSlide slide);

    /// <summary>
    /// 批量删除幻灯片
    /// </summary>
    /// <param name="indices">要删除的幻灯片索引数组 (建议降序排列)</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 移动幻灯片范围到演示文稿中的新位置
    /// 对应 SlideRange.MoveTo 方法
    /// </summary>
    /// <param name="toPos">新位置索引</param>
    void MoveTo(int toPos);

    /// <summary>
    /// 复制幻灯片范围到演示文稿中的新位置
    /// 对应 SlideRange.Duplicate 方法
    /// </summary>
    /// <returns>复制后的新幻灯片范围对象</returns>
    IPowerPointSlideRange Duplicate();
    #endregion  

    #region 内容操作 

    /// <summary>
    /// 插入新幻灯片到范围末尾
    /// </summary>
    /// <param name="layout">新幻灯片版式</param>
    /// <param name="insertIndex"></param>
    /// <returns>新插入的幻灯片对象</returns>
    IPowerPointSlide InsertNewSlide(int insertIndex, PpSlideLayout layout = PpSlideLayout.ppLayoutBlank);
    #endregion

    #region 导出和导入 (概念性)
    /// <summary>
    /// 导出幻灯片范围到 PDF 文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="overwrite">是否覆盖已存在文件</param>
    /// <returns>是否导出成功</returns>
    bool ExportToPDF(string filename, bool overwrite = true);

    /// <summary>
    /// 导出幻灯片范围到图片文件 (每张幻灯片一张图)
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="format">图片格式 (例如 "png", "jpg")</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的幻灯片数量</returns>
    int ExportToImages(string folderPath, string format = "png", string prefix = "slide_");

    #endregion    
}

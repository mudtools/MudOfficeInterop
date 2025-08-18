//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Comments 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Comments 的安全访问和操作
/// </summary>
public interface IExcelComments : IEnumerable<IExcelComment>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取评论集合中的评论数量
    /// 对应 Comments.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的评论对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">评论索引（从1开始）</param>
    /// <returns>评论对象</returns>
    IExcelComment this[int index] { get; }

    /// <summary>
    /// 获取评论集合所在的父对象（通常是工作表）
    /// 对应 Comments.Parent 属性
    /// </summary>
    object Parent { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向指定区域添加新的评论
    /// </summary>
    /// <param name="range">要添加评论的区域</param>
    /// <param name="text">评论文本内容</param>
    /// <param name="author">评论作者</param>
    /// <returns>新创建的评论对象</returns>
    IExcelComment? Add(IExcelRange range, string text, string author = "");

    /// <summary>
    /// 批量添加评论
    /// </summary>
    /// <param name="commentsData">评论数据数组，每个元素包含区域、文本和作者信息</param>
    /// <returns>成功添加的评论数量</returns>
    int AddRange(CommentData[] commentsData);

    /// <summary>
    /// 从字典添加评论
    /// </summary>
    /// <param name="cellAddress">单元格地址</param>
    /// <param name="commentText">评论文本</param>
    /// <returns>新创建的评论对象</returns>
    IExcelComment AddFromDictionary(string cellAddress, string commentText);

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据作者查找评论
    /// </summary>
    /// <param name="author">作者名称</param>
    /// <returns>匹配的评论数组</returns>
    IExcelComment[] FindByAuthor(string author);

    /// <summary>
    /// 根据文本内容查找评论
    /// </summary>
    /// <param name="text">文本内容</param>
    /// <returns>匹配的评论数组</returns>
    IExcelComment[] FindByText(string text);

    /// <summary>
    /// 根据区域查找评论
    /// </summary>
    /// <param name="range">查找区域</param>
    /// <returns>匹配的评论数组</returns>
    IExcelComment[] FindByRange(IExcelRange range);

    /// <summary>
    /// 获取指定区域内的所有评论
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的评论数组</returns>
    IExcelComment[] GetCommentsInRange(IExcelRange range);

    /// <summary>
    /// 获取可见的评论
    /// </summary>
    /// <returns>可见评论数组</returns>
    IExcelComment[] GetVisibleComments();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有评论
    /// 对应 Comments.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的评论
    /// </summary>
    /// <param name="index">要删除的评论索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的评论
    /// </summary>
    /// <param name="comment">要删除的评论对象</param>
    void Delete(IExcelComment comment);

    /// <summary>
    /// 批量删除评论
    /// </summary>
    /// <param name="indices">要删除的评论索引数组</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 显示所有评论
    /// </summary>
    void ShowAll();

    /// <summary>
    /// 隐藏所有评论
    /// </summary>
    void HideAll();

    /// <summary>
    /// 切换所有评论的可见性
    /// </summary>
    void ToggleVisibility();

    /// <summary>
    /// 刷新评论显示
    /// </summary>
    void Refresh();

    #endregion

    #region 导出和导入

    /// <summary>
    /// 导出所有评论到文本文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="includeAuthor">是否包含作者信息</param>
    /// <param name="includeAddress">是否包含单元格地址</param>
    void ExportToText(string filename, bool includeAuthor = true, bool includeAddress = true);

    /// <summary>
    /// 从文本文件导入评论
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的评论数量</returns>
    int ImportFromText(string filename);

    /// <summary>
    /// 导出所有评论到XML格式
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    void ExportToXml(string filename);

    /// <summary>
    /// 获取所有评论的文本内容
    /// </summary>
    /// <returns>评论文本数组</returns>
    string[] GetAllText();

    /// <summary>
    /// 获取所有评论的详细信息
    /// </summary>
    /// <returns>评论详细信息数组</returns>
    CommentInfo[] GetAllCommentInfo();

    #endregion

    #region 统计和分析

    /// <summary>
    /// 获取评论统计信息
    /// </summary>
    /// <returns>评论统计信息对象</returns>
    CommentStatistics GetStatistics();

    /// <summary>
    /// 获取作者统计信息
    /// </summary>
    /// <returns>作者统计信息数组</returns>
    AuthorStatistics[] GetAuthorStatistics();

    /// <summary>
    /// 获取最常见的评论文本
    /// </summary>
    /// <param name="count">返回的数量</param>
    /// <returns>最常见的评论文本数组</returns>
    string[] GetMostCommonComments(int count = 10);

    /// <summary>
    /// 获取评论长度统计
    /// </summary>
    /// <returns>长度统计信息</returns>
    LengthStatistics GetLengthStatistics();

    #endregion
}

/// <summary>
/// 评论数据结构
/// </summary>
public class CommentData
{
    /// <summary>
    /// 单元格区域
    /// </summary>
    public IExcelRange Range { get; set; }

    /// <summary>
    /// 评论文本
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// 评论作者
    /// </summary>
    public string Author { get; set; }
}

/// <summary>
/// 评论详细信息结构
/// </summary>
public class CommentInfo
{
    /// <summary>
    /// 评论索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 单元格地址
    /// </summary>
    public string Address { get; set; }

    /// <summary>
    /// 评论文本
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// 作者
    /// </summary>
    public string Author { get; set; }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool Visible { get; set; }

    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime Created { get; set; }

    /// <summary>
    /// 文本长度
    /// </summary>
    public int TextLength { get; set; }
}

/// <summary>
/// 评论统计信息结构
/// </summary>
public class CommentStatistics
{
    /// <summary>
    /// 总评论数
    /// </summary>
    public int TotalCount { get; set; }

    /// <summary>
    /// 可见评论数
    /// </summary>
    public int VisibleCount { get; set; }

    /// <summary>
    /// 隐藏评论数
    /// </summary>
    public int HiddenCount { get; set; }

    /// <summary>
    /// 平均文本长度
    /// </summary>
    public double AverageLength { get; set; }

    /// <summary>
    /// 最大文本长度
    /// </summary>
    public int MaxLength { get; set; }

    /// <summary>
    /// 最小文本长度
    /// </summary>
    public int MinLength { get; set; }

    /// <summary>
    /// 唯一作者数
    /// </summary>
    public int UniqueAuthors { get; set; }
}

/// <summary>
/// 作者统计信息结构
/// </summary>
public class AuthorStatistics
{
    /// <summary>
    /// 作者名称
    /// </summary>
    public string Author { get; set; }

    /// <summary>
    /// 评论数量
    /// </summary>
    public int CommentCount { get; set; }

    /// <summary>
    /// 平均文本长度
    /// </summary>
    public double AverageLength { get; set; }

    /// <summary>
    /// 最后评论时间
    /// </summary>
    public DateTime LastComment { get; set; }
}

/// <summary>
/// 长度统计信息结构
/// </summary>
public class LengthStatistics
{
    /// <summary>
    /// 短评论数量（1-50字符）
    /// </summary>
    public int ShortComments { get; set; }

    /// <summary>
    /// 中等评论数量（51-200字符）
    /// </summary>
    public int MediumComments { get; set; }

    /// <summary>
    /// 长评论数量（201-500字符）
    /// </summary>
    public int LongComments { get; set; }

    /// <summary>
    /// 超长评论数量（500+字符）
    /// </summary>
    public int VeryLongComments { get; set; }
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Comments 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Comments 对象的安全访问和资源管理
/// </summary>
internal class ExcelComments : IExcelComments
{
    /// <summary>
    /// 底层的 COM Comments 集合对象
    /// </summary>
    private MsExcel.Comments _comments;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelComments 实例
    /// </summary>
    /// <param name="comments">底层的 COM Comments 集合对象</param>
    internal ExcelComments(MsExcel.Comments comments)
    {
        _comments = comments;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放所有子评论对象
                for (int i = 1; i <= Count; i++)
                {
                    var comment = this[i] as ExcelComment;
                    comment?.Dispose();
                }

                // 释放底层COM对象
                if (_comments != null)
                    Marshal.ReleaseComObject(_comments);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _comments = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取评论集合中的评论数量
    /// </summary>
    public int Count => _comments?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的评论对象
    /// </summary>
    /// <param name="index">评论索引（从1开始）</param>
    /// <returns>评论对象</returns>
    public IExcelComment this[int index]
    {
        get
        {
            if (_comments == null || index < 1 || index > Count)
                return null;

            try
            {
                var comment = _comments[index] as MsExcel.Comment;
                return comment != null ? new ExcelComment(comment) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取评论集合所在的父对象
    /// </summary>
    public object Parent => _comments?.Parent;

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向指定区域添加新的评论
    /// </summary>
    /// <param name="range">要添加评论的区域</param>
    /// <param name="text">评论文本内容</param>
    /// <param name="author">评论作者</param>
    /// <returns>新创建的评论对象</returns>
    public IExcelComment? Add(IExcelRange range, string text, string author = "")
    {
        if (_comments == null || range == null || string.IsNullOrEmpty(text))
            return null;

        try
        {
            var excelRange = range as ExcelRange;
            var comment = excelRange?.InternalRange?.AddComment(text);
            if (comment != null && !string.IsNullOrEmpty(author))
            {
                // 注意：Excel中无法直接设置评论作者，作者通常是当前用户
            }

            return comment != null ? new ExcelComment(comment) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 批量添加评论
    /// </summary>
    /// <param name="commentsData">评论数据数组，每个元素包含区域、文本和作者信息</param>
    /// <returns>成功添加的评论数量</returns>
    public int AddRange(CommentData[] commentsData)
    {
        if (_comments == null || commentsData == null || commentsData.Length == 0)
            return 0;

        int successCount = 0;
        foreach (var data in commentsData)
        {
            if (Add(data.Range, data.Text, data.Author) != null)
                successCount++;
        }
        return successCount;
    }

    /// <summary>
    /// 从字典添加评论
    /// </summary>
    /// <param name="cellAddress">单元格地址</param>
    /// <param name="commentText">评论文本</param>
    /// <returns>新创建的评论对象</returns>
    public IExcelComment AddFromDictionary(string cellAddress, string commentText)
    {
        if (_comments == null || string.IsNullOrEmpty(cellAddress) || string.IsNullOrEmpty(commentText))
            return null;

        try
        {
            // 这个方法需要访问父工作表来获取区域
            var parentSheet = _comments?.Parent as MsExcel.Worksheet;
            if (parentSheet == null)
                return null;

            var range = parentSheet.Range[cellAddress] as MsExcel.Range;
            if (range == null)
                return null;

            var comment = range.AddComment(commentText) as MsExcel.Comment;
            return comment != null ? new ExcelComment(comment) : null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据作者查找评论
    /// </summary>
    /// <param name="author">作者名称</param>
    /// <returns>匹配的评论数组</returns>
    public IExcelComment[] FindByAuthor(string author)
    {
        if (_comments == null || string.IsNullOrEmpty(author) || Count == 0)
            return new IExcelComment[0];

        var result = new System.Collections.Generic.List<IExcelComment>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && comment.Author?.Contains(author) == true)
                {
                    result.Add(comment);
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据文本内容查找评论
    /// </summary>
    /// <param name="text">文本内容</param>
    /// <returns>匹配的评论数组</returns>
    public IExcelComment[] FindByText(string text)
    {
        if (_comments == null || string.IsNullOrEmpty(text) || Count == 0)
            return new IExcelComment[0];

        var result = new System.Collections.Generic.List<IExcelComment>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && comment.Text()?.Contains(text) == true)
                {
                    result.Add(comment);
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据区域查找评论
    /// </summary>
    /// <param name="range">查找区域</param>
    /// <returns>匹配的评论数组</returns>
    public IExcelComment[] FindByRange(IExcelRange range)
    {
        if (_comments == null || range == null || Count == 0)
            return new IExcelComment[0];

        var result = new System.Collections.Generic.List<IExcelComment>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                // 注意：Comment对象没有直接的区域信息，需要通过Parent获取
                // 这里提供一个简化实现
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取指定区域内的所有评论
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的评论数组</returns>
    public IExcelComment[] GetCommentsInRange(IExcelRange range)
    {
        if (_comments == null || range == null || Count == 0)
            return new IExcelComment[0];

        var result = new System.Collections.Generic.List<IExcelComment>();
        // 注意：Excel Comments集合不直接支持区域筛选
        // 这里返回所有评论作为示例
        for (int i = 1; i <= Count; i++)
        {
            var comment = this[i];
            if (comment != null)
                result.Add(comment);
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取可见的评论
    /// </summary>
    /// <returns>可见评论数组</returns>
    public IExcelComment[] GetVisibleComments()
    {
        if (_comments == null || Count == 0)
            return new IExcelComment[0];

        var result = new System.Collections.Generic.List<IExcelComment>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && comment.Visible)
                {
                    result.Add(comment);
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有评论
    /// </summary>
    public void Clear()
    {
        if (_comments == null) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    _comments[i].Delete();
                }
                catch
                {
                    // 忽略删除过程中的异常
                }
            }
        }
        catch
        {
            // 忽略清空过程中的异常
        }
    }

    /// <summary>
    /// 删除指定索引的评论
    /// </summary>
    /// <param name="index">要删除的评论索引</param>
    public void Delete(int index)
    {
        if (_comments == null || index < 1 || index > Count)
            return;

        try
        {
            _comments[index].Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的评论
    /// </summary>
    /// <param name="comment">要删除的评论对象</param>
    public void Delete(IExcelComment comment)
    {
        if (_comments == null || comment == null)
            return;

        try
        {
            comment.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 批量删除评论
    /// </summary>
    /// <param name="indices">要删除的评论索引数组</param>
    public void DeleteRange(int[] indices)
    {
        if (_comments == null || indices == null || indices.Length == 0)
            return;

        // 按降序排列索引，避免删除时索引变化
        Array.Sort(indices, (a, b) => b.CompareTo(a));

        foreach (int index in indices)
        {
            Delete(index);
        }
    }

    /// <summary>
    /// 显示所有评论
    /// </summary>
    public void ShowAll()
    {
        if (_comments == null || Count == 0)
            return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                    comment.Visible = true;
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
    }

    /// <summary>
    /// 隐藏所有评论
    /// </summary>
    public void HideAll()
    {
        if (_comments == null || Count == 0)
            return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                    comment.Visible = false;
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
    }

    /// <summary>
    /// 切换所有评论的可见性
    /// </summary>
    public void ToggleVisibility()
    {
        if (_comments == null || Count == 0)
            return;

        bool hasVisible = false;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && comment.Visible)
                {
                    hasVisible = true;
                    break;
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }

        // 切换所有评论的可见性
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                    comment.Visible = !hasVisible;
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
    }

    /// <summary>
    /// 刷新评论显示
    /// </summary>
    public void Refresh()
    {
        // Excel Comments通常会自动刷新
        // 这里提供一个空实现以保持接口一致性
    }

    #endregion

    #region 导出和导入

    /// <summary>
    /// 导出所有评论到文本文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="includeAuthor">是否包含作者信息</param>
    /// <param name="includeAddress">是否包含单元格地址</param>
    public void ExportToText(string filename, bool includeAuthor = true, bool includeAddress = true)
    {
        if (_comments == null || string.IsNullOrEmpty(filename) || Count == 0)
            return;

        try
        {
            using (var writer = new System.IO.StreamWriter(filename, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Excel Comments Export");
                writer.WriteLine("====================");
                writer.WriteLine($"Export Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine($"Total Comments: {Count}");
                writer.WriteLine();

                for (int i = 1; i <= Count; i++)
                {
                    try
                    {
                        var comment = this[i];
                        if (comment != null)
                        {
                            writer.WriteLine($"Comment #{i}");

                            if (includeAddress)
                                writer.WriteLine($"Address: {comment.Parent?.Address ?? "N/A"}");

                            if (includeAuthor)
                                writer.WriteLine($"Author: {comment.Author ?? "N/A"}");

                            writer.WriteLine($"Visible: {comment.Visible}");
                            writer.WriteLine($"Text: {comment.Text() ?? "N/A"}");
                            writer.WriteLine(new string('-', 40));
                            writer.WriteLine();
                        }
                    }
                    catch
                    {
                        // 忽略单个评论导出异常
                    }
                }
            }
        }
        catch
        {
            // 忽略文件操作异常
        }
    }

    /// <summary>
    /// 从文本文件导入评论
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的评论数量</returns>
    public int ImportFromText(string filename)
    {
        if (_comments == null || string.IsNullOrEmpty(filename))
            return 0;

        // 注意：Excel Comments不支持直接导入
        // 这里提供一个示例实现框架
        return 0;
    }

    /// <summary>
    /// 导出所有评论到XML格式
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    public void ExportToXml(string filename)
    {
        if (_comments == null || string.IsNullOrEmpty(filename) || Count == 0)
            return;

        try
        {
            var xmlDoc = new System.Xml.XmlDocument();
            var root = xmlDoc.CreateElement("Comments");
            xmlDoc.AppendChild(root);

            var exportDate = xmlDoc.CreateElement("ExportDate");
            exportDate.InnerText = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            root.AppendChild(exportDate);

            var totalCount = xmlDoc.CreateElement("TotalCount");
            totalCount.InnerText = Count.ToString();
            root.AppendChild(totalCount);

            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var comment = this[i];
                    if (comment != null)
                    {
                        var commentElement = xmlDoc.CreateElement("Comment");

                        var indexElement = xmlDoc.CreateElement("Index");
                        indexElement.InnerText = i.ToString();
                        commentElement.AppendChild(indexElement);

                        var addressElement = xmlDoc.CreateElement("Address");
                        addressElement.InnerText = comment.Parent?.Address ?? "N/A";
                        commentElement.AppendChild(addressElement);

                        var authorElement = xmlDoc.CreateElement("Author");
                        authorElement.InnerText = comment.Author ?? "N/A";
                        commentElement.AppendChild(authorElement);

                        var visibleElement = xmlDoc.CreateElement("Visible");
                        visibleElement.InnerText = comment.Visible.ToString();
                        commentElement.AppendChild(visibleElement);

                        var textElement = xmlDoc.CreateElement("Text");
                        textElement.InnerText = comment.Text() ?? "N/A";
                        commentElement.AppendChild(textElement);

                        root.AppendChild(commentElement);
                    }
                }
                catch
                {
                    // 忽略单个评论导出异常
                }
            }

            xmlDoc.Save(filename);
        }
        catch
        {
            // 忽略文件操作异常
        }
    }

    /// <summary>
    /// 获取所有评论的文本内容
    /// </summary>
    /// <returns>评论文本数组</returns>
    public string[] GetAllText()
    {
        if (_comments == null || Count == 0)
            return new string[0];

        var result = new System.Collections.Generic.List<string>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && !string.IsNullOrEmpty(comment.Text()))
                {
                    result.Add(comment.Text());
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取所有评论的详细信息
    /// </summary>
    /// <returns>评论详细信息数组</returns>
    public CommentInfo[] GetAllCommentInfo()
    {
        if (_comments == null || Count == 0)
            return new CommentInfo[0];

        var result = new System.Collections.Generic.List<CommentInfo>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                {
                    var info = new CommentInfo
                    {
                        Index = i,
                        Address = comment.Parent?.Address ?? "N/A",
                        Text = comment.Text() ?? "",
                        Author = comment.Author ?? "N/A",
                        Visible = comment.Visible,
                        Created = DateTime.Now, // Excel不提供创建时间
                        TextLength = comment.Text()?.Length ?? 0
                    };
                    result.Add(info);
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 统计和分析

    /// <summary>
    /// 获取评论统计信息
    /// </summary>
    /// <returns>评论统计信息对象</returns>
    public CommentStatistics GetStatistics()
    {
        var stats = new CommentStatistics
        {
            TotalCount = Count,
            VisibleCount = 0,
            HiddenCount = 0,
            AverageLength = 0,
            MaxLength = 0,
            MinLength = int.MaxValue,
            UniqueAuthors = 0
        };

        if (_comments == null || Count == 0)
            return stats;

        int totalLength = 0;
        var authors = new System.Collections.Generic.HashSet<string>();

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                {
                    if (comment.Visible)
                        stats.VisibleCount++;
                    else
                        stats.HiddenCount++;

                    int textLength = comment.Text()?.Length ?? 0;
                    totalLength += textLength;

                    if (textLength > stats.MaxLength)
                        stats.MaxLength = textLength;

                    if (textLength < stats.MinLength)
                        stats.MinLength = textLength;

                    if (!string.IsNullOrEmpty(comment.Author))
                        authors.Add(comment.Author);
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }

        stats.MinLength = stats.MinLength == int.MaxValue ? 0 : stats.MinLength;
        stats.AverageLength = Count > 0 ? (double)totalLength / Count : 0;
        stats.UniqueAuthors = authors.Count;

        return stats;
    }

    /// <summary>
    /// 获取作者统计信息
    /// </summary>
    /// <returns>作者统计信息数组</returns>
    public AuthorStatistics[] GetAuthorStatistics()
    {
        if (_comments == null || Count == 0)
            return new AuthorStatistics[0];

        var authorStats = new System.Collections.Generic.Dictionary<string, AuthorStatistics>();

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                {
                    string author = comment.Author ?? "Unknown";

                    if (!authorStats.ContainsKey(author))
                    {
                        authorStats[author] = new AuthorStatistics
                        {
                            Author = author,
                            CommentCount = 0,
                            AverageLength = 0,
                            LastComment = DateTime.MinValue
                        };
                    }

                    var stats = authorStats[author];
                    stats.CommentCount++;
                    stats.AverageLength = ((stats.AverageLength * (stats.CommentCount - 1)) +
                                         (comment.Text()?.Length ?? 0)) / stats.CommentCount;
                    stats.LastComment = DateTime.Now; // Excel不提供实际时间
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }

        return new System.Collections.Generic.List<AuthorStatistics>(authorStats.Values).ToArray();
    }

    /// <summary>
    /// 获取最常见的评论文本
    /// </summary>
    /// <param name="count">返回的数量</param>
    /// <returns>最常见的评论文本数组</returns>
    public string[] GetMostCommonComments(int count = 10)
    {
        if (_comments == null || Count == 0 || count <= 0)
            return new string[0];

        var textCount = new System.Collections.Generic.Dictionary<string, int>();

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null && !string.IsNullOrEmpty(comment.Text()))
                {
                    string text = comment.Text();
                    if (textCount.ContainsKey(text))
                        textCount[text]++;
                    else
                        textCount[text] = 1;
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }

        // 按出现次数排序并返回前N个
        var sorted = textCount.OrderByDescending(x => x.Value)
                             .Take(count)
                             .Select(x => x.Key)
                             .ToArray();

        return sorted;
    }

    /// <summary>
    /// 获取评论长度统计
    /// </summary>
    /// <returns>长度统计信息</returns>
    public LengthStatistics GetLengthStatistics()
    {
        var stats = new LengthStatistics
        {
            ShortComments = 0,
            MediumComments = 0,
            LongComments = 0,
            VeryLongComments = 0
        };

        if (_comments == null || Count == 0)
            return stats;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var comment = this[i];
                if (comment != null)
                {
                    int length = comment.Text()?.Length ?? 0;

                    if (length <= 50)
                        stats.ShortComments++;
                    else if (length <= 200)
                        stats.MediumComments++;
                    else if (length <= 500)
                        stats.LongComments++;
                    else
                        stats.VeryLongComments++;
                }
            }
            catch
            {
                // 忽略单个评论访问异常
            }
        }

        return stats;
    }

    public IEnumerator<IExcelComment> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return new ExcelComment(_comments[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
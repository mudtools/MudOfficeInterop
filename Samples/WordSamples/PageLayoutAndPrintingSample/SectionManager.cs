//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace PageLayoutAndPrintingSample
{
    /// <summary>
    /// 节管理器类
    /// </summary>
    public class SectionManager
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public SectionManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 在指定位置插入分页符
        /// </summary>
        /// <param name="position">插入位置</param>
        public void InsertPageBreak(IWordRange position)
        {
            try
            {
                position.InsertBreak(WdBreakType.wdPageBreak);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入分页符时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在文档末尾插入分页符
        /// </summary>
        public void InsertPageBreakAtEnd()
        {
            try
            {
                using var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                InsertPageBreak(range);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"在文档末尾插入分页符时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在指定位置插入分节符（下一页）
        /// </summary>
        /// <param name="position">插入位置</param>
        public void InsertSectionBreakNextPage(IWordRange position)
        {
            try
            {
                position.InsertBreak(WdBreakType.wdSectionBreakNextPage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入分节符（下一页）时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在文档末尾插入分节符（下一页）
        /// </summary>
        public void InsertSectionBreakNextPageAtEnd()
        {
            try
            {
                using var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                InsertSectionBreakNextPage(range);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"在文档末尾插入分节符（下一页）时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在指定位置插入分节符（连续）
        /// </summary>
        /// <param name="position">插入位置</param>
        public void InsertSectionBreakContinuous(IWordRange position)
        {
            try
            {
                position.InsertBreak(WdBreakType.wdSectionBreakContinuous);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入分节符（连续）时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在指定位置插入分节符（偶数页）
        /// </summary>
        /// <param name="position">插入位置</param>
        public void InsertSectionBreakEvenPage(IWordRange position)
        {
            try
            {
                position.InsertBreak(WdBreakType.wdSectionBreakEvenPage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入分节符（偶数页）时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在指定位置插入分节符（奇数页）
        /// </summary>
        /// <param name="position">插入位置</param>
        public void InsertSectionBreakOddPage(IWordRange position)
        {
            try
            {
                position.InsertBreak(WdBreakType.wdSectionBreakOddPage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入分节符（奇数页）时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取节的数量
        /// </summary>
        /// <returns>节的数量</returns>
        public int GetSectionCount()
        {
            try
            {
                return _document.Sections.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取节数量时出错: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 获取指定节的起始和结束范围
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <returns>节范围信息</returns>
        public string GetSectionRangeInfo(int sectionIndex)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var range = section.Range;

                    return $"第 {sectionIndex} 节范围: 起始位置 {range.Start}, 结束位置 {range.End}";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取节范围信息时出错: {ex.Message}");
            }

            return $"无法获取第 {sectionIndex} 节的范围信息";
        }

        /// <summary>
        /// 为指定节添加内容
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="content">内容文本</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        public void AddContentToSection(int sectionIndex, string content, string fontName = "宋体", float fontSize = 12)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var range = section.Range;

                    // 将光标移到节的末尾
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加内容
                    range.Text = content;
                    range.Font.Name = fontName;
                    range.Font.Size = fontSize;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"为节添加内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 复制节的内容到新节
        /// </summary>
        /// <param name="sourceSectionIndex">源节索引</param>
        /// <param name="addSectionBreak">是否在复制前添加分节符</param>
        /// <returns>新节的索引</returns>
        public int CopySectionContent(int sourceSectionIndex, bool addSectionBreak = true)
        {
            try
            {
                if (sourceSectionIndex > 0 && sourceSectionIndex <= _document.Sections.Count)
                {
                    // 如果需要，先添加分节符
                    if (addSectionBreak)
                    {
                        InsertSectionBreakNextPageAtEnd();
                    }

                    // 获取源节内容
                    using var sourceSection = _document.Sections[sourceSectionIndex];
                    using var sourceRange = sourceSection.Range;
                    string content = sourceRange.Text;

                    // 获取新节（刚刚创建的）
                    int newSectionIndex = _document.Sections.Count;
                    var newSection = _document.Sections[newSectionIndex];
                    var newRange = newSection.Range;

                    // 清空新节内容并添加复制的内容
                    newRange.Text = content;

                    return newSectionIndex;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"复制节内容时出错: {ex.Message}");
            }

            return -1;
        }

        /// <summary>
        /// 删除指定节（通过插入连续分节符合并节）
        /// </summary>
        /// <param name="sectionIndex">要删除的节索引</param>
        /// <returns>是否删除成功</returns>
        public bool DeleteSection(int sectionIndex)
        {
            try
            {
                if (sectionIndex > 1 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var range = section.Range;

                    // 删除节的内容和分节符
                    range.Delete();

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除节时出错: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// 交换两个节的内容
        /// </summary>
        /// <param name="firstSectionIndex">第一个节索引</param>
        /// <param name="secondSectionIndex">第二个节索引</param>
        /// <returns>是否交换成功</returns>
        public bool SwapSections(int firstSectionIndex, int secondSectionIndex)
        {
            try
            {
                if (firstSectionIndex > 0 && firstSectionIndex <= _document.Sections.Count &&
                    secondSectionIndex > 0 && secondSectionIndex <= _document.Sections.Count &&
                    firstSectionIndex != secondSectionIndex)
                {
                    using var firstSection = _document.Sections[firstSectionIndex];
                    using var secondSection = _document.Sections[secondSectionIndex];

                    using var firstRange = firstSection.Range;
                    using var secondRange = secondSection.Range;

                    string firstContent = firstRange.Text;
                    string secondContent = secondRange.Text;

                    // 交换内容
                    firstRange.Text = secondContent;
                    secondRange.Text = firstContent;

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"交换节内容时出错: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// 获取所有节的信息
        /// </summary>
        /// <returns>节信息列表</returns>
        public List<string> GetAllSectionsInfo()
        {
            var sectionsInfo = new List<string>();

            try
            {
                int sectionCount = _document.Sections.Count;
                sectionsInfo.Add($"文档共有 {sectionCount} 个节");

                for (int i = 1; i <= sectionCount; i++)
                {
                    var sectionInfo = GetSectionRangeInfo(i);
                    sectionsInfo.Add(sectionInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取所有节信息时出错: {ex.Message}");
                sectionsInfo.Add("获取节信息失败");
            }

            return sectionsInfo;
        }
    }
}
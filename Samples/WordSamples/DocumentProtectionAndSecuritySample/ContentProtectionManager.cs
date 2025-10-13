//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace DocumentProtectionAndSecuritySample
{
    /// <summary>
    /// 内容保护管理器类
    /// </summary>
    public class ContentProtectionManager
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public ContentProtectionManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 添加书签保护
        /// </summary>
        /// <param name="bookmarkName">书签名称</param>
        /// <param name="range">文档范围</param>
        /// <param name="editorType">编辑者类型</param>
        /// <returns>是否添加成功</returns>
        public bool AddBookmarkProtection(string bookmarkName, IWordRange range, WdEditorType editorType = WdEditorType.wdEditorOwners)
        {
            try
            {
                // 添加书签
                var bookmark = _document.Bookmarks.Add(bookmarkName, range);

                // 为书签内容添加编辑权限
                bookmark.Range.Editors.Add(editorType);

                Console.WriteLine($"书签保护已添加: {bookmarkName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加书签保护时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加表单字段保护
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="fieldName">字段名称</param>
        /// <param name="fieldType">字段类型</param>
        /// <param name="defaultValue">默认值</param>
        /// <param name="isReadOnly">是否只读</param>
        /// <returns>表单字段对象</returns>
        public IWordFormField AddFormFieldProtection(
            IWordRange range,
            string fieldName,
            WdFieldType fieldType,
            string defaultValue = "",
            bool isReadOnly = false)
        {
            try
            {
                // 添加表单字段
                var formField = range.FormFields.Add(range, fieldType);
                formField.Name = fieldName;

                switch (fieldType)
                {
                    case WdFieldType.wdFieldFormTextInput:
                        formField.TextInput.Default = defaultValue;
                        if (isReadOnly)
                        {
                            formField.TextInput.EditType(WdTextInputType.wdRegularText, defaultValue, true);
                        }
                        break;

                    case WdFieldType.wdFieldFormCheckBox:
                        formField.CheckBox.Default = !string.IsNullOrEmpty(defaultValue) && defaultValue.ToLower() == "true";
                        break;

                    case WdFieldType.wdFieldFormDropDown:
                        // 下拉字段需要额外设置选项
                        break;
                }

                Console.WriteLine($"表单字段保护已添加: {fieldName}");
                return formField;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加表单字段保护时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 添加范围保护
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="protectionType">保护类型</param>
        /// <returns>是否添加成功</returns>
        public bool AddRangeProtection(IWordRange range, WdContentControlType protectionType)
        {
            try
            {
                // 添加内容控件
                var contentControl = range.ContentControls.Add(protectionType);

                // 设置保护属性
                contentControl.LockContentControl = true;
                contentControl.LockContents = true;

                Console.WriteLine($"范围保护已添加: {protectionType}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加范围保护时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建受保护的表单文档
        /// </summary>
        /// <param name="title">文档标题</param>
        /// <param name="formFields">表单字段定义</param>
        /// <returns>是否创建成功</returns>
        public bool CreateProtectedFormDocument(string title, List<FormFieldDefinition> formFields)
        {
            try
            {
                // 清空文档内容
                _document.Range().Text = "";

                // 添加标题
                var titleRange = _document.Range();
                titleRange.Text = $"{title}\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = 1;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加表单字段
                var contentRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);

                foreach (var field in formFields)
                {
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加字段标签
                    contentRange.Text = $"{field.Label}：";
                    contentRange.Font.Bold = 1;
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加字段占位符
                    switch (field.FieldType)
                    {
                        case WdFieldType.wdFieldFormTextInput:
                            contentRange.Text = "________________________\n";
                            break;

                        case WdFieldType.wdFieldFormCheckBox:
                            contentRange.Text = "[  ] 是    [  ] 否\n";
                            break;
                    }

                    contentRange.Font.Bold = 0;
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                Console.WriteLine("受保护的表单文档已创建");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建受保护的表单文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建机密内容区域
        /// </summary>
        /// <param name="content">机密内容</param>
        /// <param name="allowedEditors">允许的编辑者列表</param>
        /// <returns>是否创建成功</returns>
        public bool CreateConfidentialSection(string content, List<string> allowedEditors)
        {
            try
            {
                var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);

                // 添加机密内容标记
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n【机密内容开始】\n";
                range.Font.Bold = 1;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加机密内容
                range.Text = $"{content}\n";
                range.Font.Bold = 0;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加结束标记
                range.Text = "【机密内容结束】\n";
                range.Font.Bold = 1;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 为机密内容添加书签保护
                var confidentialRange = _document.Range(
                    range.Start - content.Length - 20, // 包含标记和内容
                    range.Start - 10 // 不包含结束标记
                );

                var bookmark = _document.Bookmarks.Add("ConfidentialSection", confidentialRange);

                // 设置编辑权限
                if (allowedEditors != null && allowedEditors.Any())
                {
                    foreach (var editor in allowedEditors)
                    {
                        bookmark.Range.Editors.Add(editor);
                    }
                }
                else
                {
                    // 默认只允许所有者编辑
                    bookmark.Range.Editors.Add(WdEditorType.wdEditorOwners);
                }

                Console.WriteLine("机密内容区域已创建");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建机密内容区域时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建可编辑区域
        /// </summary>
        /// <param name="content">区域内容</param>
        /// <param name="editorType">编辑者类型</param>
        /// <param name="editorName">编辑者名称</param>
        /// <returns>是否创建成功</returns>
        public bool CreateEditableSection(string content, WdEditorType editorType, string editorName = null)
        {
            try
            {
                var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);

                // 添加内容
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = $"{content}\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加可编辑区域
                var editableRange = _document.EditableRanges.Add(range);

                if (editorType == WdEditorType.wdEditorEveryone)
                {
                    editableRange.Editors.Add(WdEditorType.wdEditorEveryone);
                }
                else if (!string.IsNullOrEmpty(editorName))
                {
                    editableRange.Editors.Add(editorName);
                }

                Console.WriteLine("可编辑区域已创建");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建可编辑区域时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取所有受保护内容信息
        /// </summary>
        /// <returns>受保护内容信息列表</returns>
        public List<ProtectedContentInfo> GetAllProtectedContentInfo()
        {
            var protectedContents = new List<ProtectedContentInfo>();

            try
            {
                // 获取书签保护信息
                for (int i = 1; i <= _document.Bookmarks.Count; i++)
                {
                    var bookmark = _document.Bookmarks.Item(i);
                    var info = new ProtectedContentInfo
                    {
                        ContentType = "书签",
                        Name = bookmark.Name,
                        StartPosition = bookmark.Range.Start,
                        EndPosition = bookmark.Range.End,
                        Length = bookmark.Range.End - bookmark.Range.Start
                    };
                    protectedContents.Add(info);
                }

                // 获取表单字段保护信息
                for (int i = 1; i <= _document.FormFields.Count; i++)
                {
                    var formField = _document.FormFields.Item(i);
                    var info = new ProtectedContentInfo
                    {
                        ContentType = "表单字段",
                        Name = formField.Name,
                        StartPosition = formField.Range.Start,
                        EndPosition = formField.Range.End,
                        Length = formField.Range.End - formField.Range.Start
                    };
                    protectedContents.Add(info);
                }

                // 获取可编辑区域信息
                for (int i = 1; i <= _document.EditableRanges.Count; i++)
                {
                    var editableRange = _document.EditableRanges.Item(i);
                    var info = new ProtectedContentInfo
                    {
                        ContentType = "可编辑区域",
                        Name = $"EditableRange{i}",
                        StartPosition = editableRange.Range.Start,
                        EndPosition = editableRange.Range.End,
                        Length = editableRange.Range.End - editableRange.Range.Start
                    };
                    protectedContents.Add(info);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取受保护内容信息时出错: {ex.Message}");
            }

            return protectedContents;
        }
    }

    /// <summary>
    /// 表单字段定义类
    /// </summary>
    public class FormFieldDefinition
    {
        /// <summary>
        /// 字段标签
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 字段名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 字段类型
        /// </summary>
        public WdFieldType FieldType { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public string DefaultValue { get; set; }

        /// <summary>
        /// 是否只读
        /// </summary>
        public bool IsReadOnly { get; set; }
    }

    /// <summary>
    /// 受保护内容信息类
    /// </summary>
    public class ProtectedContentInfo
    {
        /// <summary>
        /// 内容类型
        /// </summary>
        public string ContentType { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 起始位置
        /// </summary>
        public int StartPosition { get; set; }

        /// <summary>
        /// 结束位置
        /// </summary>
        public int EndPosition { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// 生成内容信息报告
        /// </summary>
        /// <returns>信息报告</returns>
        public string GenerateReport()
        {
            return $"受保护内容信息:\n" +
                   $"  类型: {ContentType}\n" +
                   $"  名称: {Name}\n" +
                   $"  位置: {StartPosition}-{EndPosition}\n" +
                   $"  长度: {Length}";
        }
    }
}
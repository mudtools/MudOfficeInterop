//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;
using System.Linq;
using System.Text;

namespace RibbonCustomizationSample
{
    /// <summary>
    /// Ribbon管理器类
    /// </summary>
    public class RibbonManager
    {
        private readonly IWordApplication _application;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        public RibbonManager(IWordApplication application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// 创建自定义选项卡XML
        /// </summary>
        /// <param name="tabId">选项卡ID</param>
        /// <param name="tabLabel">选项卡标签</param>
        /// <param name="groups">组定义列表</param>
        /// <returns>Ribbon XML字符串</returns>
        public string CreateCustomTabXml(string tabId, string tabLabel, List<RibbonGroupDefinition> groups)
        {
            try
            {
                var xmlBuilder = new StringBuilder();
                xmlBuilder.AppendLine("<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>");
                xmlBuilder.AppendLine("  <ribbon>");
                xmlBuilder.AppendLine("    <tabs>");
                xmlBuilder.AppendLine($"      <tab id='{tabId}' label='{tabLabel}'>");

                foreach (var group in groups)
                {
                    xmlBuilder.AppendLine($"        <group id='{group.Id}' label='{group.Label}'>");

                    foreach (var control in group.Controls)
                    {
                        switch (control.ControlType)
                        {
                            case RibbonControlType.Button:
                                xmlBuilder.AppendLine($"          <button id='{control.Id}' label='{control.Label}' " +
                                    $"onAction='{control.OnAction}' size='{control.Size}' " +
                                    $"imageMso='{control.ImageMso}' />");
                                break;

                            case RibbonControlType.ToggleButton:
                                xmlBuilder.AppendLine($"          <toggleButton id='{control.Id}' label='{control.Label}' " +
                                    $"onAction='{control.OnAction}' getPressed='{control.GetPressed}' " +
                                    $"imageMso='{control.ImageMso}' />");
                                break;

                            case RibbonControlType.DropDown:
                                xmlBuilder.AppendLine($"          <dropDown id='{control.Id}' label='{control.Label}' " +
                                    $"onAction='{control.OnAction}' {control.AdditionalAttributes} />");
                                break;

                            case RibbonControlType.EditBox:
                                xmlBuilder.AppendLine($"          <editBox id='{control.Id}' label='{control.Label}' " +
                                    $"onChange='{control.OnChange}' {control.AdditionalAttributes} />");
                                break;

                            case RibbonControlType.Separator:
                                xmlBuilder.AppendLine("          <separator />");
                                break;
                        }
                    }

                    xmlBuilder.AppendLine("        </group>");
                }

                xmlBuilder.AppendLine("      </tab>");
                xmlBuilder.AppendLine("    </tabs>");
                xmlBuilder.AppendLine("  </ribbon>");
                xmlBuilder.AppendLine("</customUI>");

                Console.WriteLine("自定义选项卡XML已创建");
                return xmlBuilder.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建自定义选项卡XML时出错: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 创建文档工具选项卡
        /// </summary>
        /// <returns>Ribbon XML字符串</returns>
        public string CreateDocumentToolsTab()
        {
            try
            {
                var groups = new List<RibbonGroupDefinition>
                {
                    new RibbonGroupDefinition
                    {
                        Id = "grpFormatting",
                        Label = "格式化工具",
                        Controls = new List<RibbonControlDefinition>
                        {
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnHeading1",
                                Label = "标题1",
                                Size = "large",
                                OnAction = "OnHeading1",
                                ImageMso = "StyleHeading1"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnHeading2",
                                Label = "标题2",
                                Size = "large",
                                OnAction = "OnHeading2",
                                ImageMso = "StyleHeading2"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Separator
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnBold",
                                Label = "加粗",
                                OnAction = "OnBold",
                                ImageMso = "Bold"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnItalic",
                                Label = "斜体",
                                OnAction = "OnItalic",
                                ImageMso = "Italic"
                            }
                        }
                    },
                    new RibbonGroupDefinition
                    {
                        Id = "grpTables",
                        Label = "表格工具",
                        Controls = new List<RibbonControlDefinition>
                        {
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnInsertTable",
                                Label = "插入表格",
                                Size = "large",
                                OnAction = "OnInsertTable",
                                ImageMso = "TableInsertTable"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnFormatTable",
                                Label = "格式化表格",
                                Size = "large",
                                OnAction = "OnFormatTable",
                                ImageMso = "TableStyles"
                            }
                        }
                    },
                    new RibbonGroupDefinition
                    {
                        Id = "grpAutomation",
                        Label = "自动化",
                        Controls = new List<RibbonControlDefinition>
                        {
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnAutoNumber",
                                Label = "自动编号",
                                OnAction = "OnAutoNumber"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnGenerateTOC",
                                Label = "生成目录",
                                OnAction = "OnGenerateTOC"
                            }
                        }
                    }
                };

                return CreateCustomTabXml("tabDocumentTools", "文档工具", groups);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建文档工具选项卡时出错: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 创建动态工具选项卡
        /// </summary>
        /// <returns>Ribbon XML字符串</returns>
        public string CreateDynamicToolsTab()
        {
            try
            {
                var groups = new List<RibbonGroupDefinition>
                {
                    new RibbonGroupDefinition
                    {
                        Id = "grpSelection",
                        Label = "选择操作",
                        Controls = new List<RibbonControlDefinition>
                        {
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnCopy",
                                Label = "复制",
                                OnAction = "OnCopy",
                                GetEnabled = "IsTextSelected",
                                ImageMso = "Copy"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnCut",
                                Label = "剪切",
                                OnAction = "OnCut",
                                GetEnabled = "IsTextSelected",
                                ImageMso = "Cut"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnPaste",
                                Label = "粘贴",
                                OnAction = "OnPaste",
                                GetEnabled = "IsClipboardNotEmpty",
                                ImageMso = "Paste"
                            }
                        }
                    },
                    new RibbonGroupDefinition
                    {
                        Id = "grpDocument",
                        Label = "文档状态",
                        Controls = new List<RibbonControlDefinition>
                        {
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnSave",
                                Label = "保存",
                                OnAction = "OnSave",
                                GetEnabled = "IsDocumentModified",
                                ImageMso = "FileSave"
                            },
                            new RibbonControlDefinition
                            {
                                ControlType = RibbonControlType.Button,
                                Id = "btnPrint",
                                Label = "打印",
                                OnAction = "OnPrint",
                                ImageMso = "Print"
                            }
                        }
                    }
                };

                return CreateCustomTabXml("tabDynamicTools", "动态工具", groups);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建动态工具选项卡时出错: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 验证Ribbon XML格式
        /// </summary>
        /// <param name="xml">XML字符串</param>
        /// <returns>是否格式正确</returns>
        public bool ValidateRibbonXml(string xml)
        {
            try
            {
                // 简单验证XML是否包含必要的元素
                bool hasCustomUI = xml.Contains("<customUI");
                bool hasRibbon = xml.Contains("<ribbon>");
                bool hasTabs = xml.Contains("<tabs>");

                bool isValid = hasCustomUI && hasRibbon && hasTabs;

                Console.WriteLine($"Ribbon XML验证结果: {(isValid ? "有效" : "无效")}");
                return isValid;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"验证Ribbon XML时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成Ribbon回调函数模板
        /// </summary>
        /// <param name="controls">控件定义列表</param>
        /// <returns>回调函数模板代码</returns>
        public string GenerateCallbackTemplates(List<RibbonControlDefinition> controls)
        {
            try
            {
                var codeBuilder = new StringBuilder();
                codeBuilder.AppendLine("// === Ribbon回调函数模板 ===");
                codeBuilder.AppendLine();

                // 生成OnAction回调函数
                var actionControls = controls.Where(c => !string.IsNullOrEmpty(c.OnAction)).GroupBy(c => c.OnAction).Select(g => g.First());
                foreach (var control in actionControls)
                {
                    codeBuilder.AppendLine($"public void {control.OnAction}(IRibbonControl control)");
                    codeBuilder.AppendLine("{");
                    codeBuilder.AppendLine("    // TODO: 实现功能逻辑");
                    codeBuilder.AppendLine($"    System.Windows.Forms.MessageBox.Show(\"执行 {control.Label} 功能\");");
                    codeBuilder.AppendLine("}");
                    codeBuilder.AppendLine();
                }

                // 生成GetEnabled回调函数
                var enabledControls = controls.Where(c => !string.IsNullOrEmpty(c.GetEnabled)).GroupBy(c => c.GetEnabled).Select(g => g.First());
                foreach (var control in enabledControls)
                {
                    codeBuilder.AppendLine($"public bool {control.GetEnabled}(IRibbonControl control)");
                    codeBuilder.AppendLine("{");
                    codeBuilder.AppendLine("    // TODO: 实现启用状态逻辑");
                    codeBuilder.AppendLine("    return true;");
                    codeBuilder.AppendLine("}");
                    codeBuilder.AppendLine();
                }

                // 生成GetPressed回调函数
                var pressedControls = controls.Where(c => !string.IsNullOrEmpty(c.GetPressed)).GroupBy(c => c.GetPressed).Select(g => g.First());
                foreach (var control in pressedControls)
                {
                    codeBuilder.AppendLine($"public bool {control.GetPressed}(IRibbonControl control)");
                    codeBuilder.AppendLine("{");
                    codeBuilder.AppendLine("    // TODO: 实现按下状态逻辑");
                    codeBuilder.AppendLine("    return false;");
                    codeBuilder.AppendLine("}");
                    codeBuilder.AppendLine();
                }

                Console.WriteLine("Ribbon回调函数模板已生成");
                return codeBuilder.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成回调函数模板时出错: {ex.Message}");
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// Ribbon组定义类
    /// </summary>
    public class RibbonGroupDefinition
    {
        /// <summary>
        /// 组ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 组标签
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 控件定义列表
        /// </summary>
        public List<RibbonControlDefinition> Controls { get; set; } = new List<RibbonControlDefinition>();
    }

    /// <summary>
    /// Ribbon控件定义类
    /// </summary>
    public class RibbonControlDefinition
    {
        /// <summary>
        /// 控件类型
        /// </summary>
        public RibbonControlType ControlType { get; set; }

        /// <summary>
        /// 控件ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 控件标签
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 控件大小
        /// </summary>
        public string Size { get; set; } = "normal";

        /// <summary>
        /// 点击事件回调函数
        /// </summary>
        public string OnAction { get; set; }

        /// <summary>
        /// 启用状态回调函数
        /// </summary>
        public string GetEnabled { get; set; }

        /// <summary>
        /// 按下状态回调函数
        /// </summary>
        public string GetPressed { get; set; }

        /// <summary>
        /// 变化事件回调函数
        /// </summary>
        public string OnChange { get; set; }

        /// <summary>
        /// 图像标识
        /// </summary>
        public string ImageMso { get; set; }

        /// <summary>
        /// 附加属性
        /// </summary>
        public string AdditionalAttributes { get; set; }
    }

    /// <summary>
    /// Ribbon控件类型枚举
    /// </summary>
    public enum RibbonControlType
    {
        /// <summary>
        /// 按钮
        /// </summary>
        Button,

        /// <summary>
        /// 切换按钮
        /// </summary>
        ToggleButton,

        /// <summary>
        /// 下拉列表
        /// </summary>
        DropDown,

        /// <summary>
        /// 编辑框
        /// </summary>
        EditBox,

        /// <summary>
        /// 分隔符
        /// </summary>
        Separator
    }
}
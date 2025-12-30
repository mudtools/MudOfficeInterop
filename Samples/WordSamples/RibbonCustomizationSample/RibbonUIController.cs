//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System.Text;

namespace RibbonCustomizationSample
{
    /// <summary>
    /// Ribbon UI控制器类
    /// </summary>
    public class RibbonUIController
    {
        private readonly IWordApplication _application;
        private readonly RibbonManager _ribbonManager;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        public RibbonUIController(IWordApplication? Application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _ribbonManager = new RibbonManager(application);
        }

        /// <summary>
        /// 创建文档工具Ribbon界面
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateDocumentToolsRibbon()
        {
            try
            {
                Console.WriteLine("开始创建文档工具Ribbon界面...");

                // 创建文档工具选项卡XML
                string ribbonXml = _ribbonManager.CreateDocumentToolsTab();

                if (string.IsNullOrEmpty(ribbonXml))
                {
                    Console.WriteLine("创建文档工具选项卡XML失败");
                    return false;
                }

                // 验证XML
                bool isValid = _ribbonManager.ValidateRibbonXml(ribbonXml);
                if (!isValid)
                {
                    Console.WriteLine("文档工具选项卡XML验证失败");
                    return false;
                }

                Console.WriteLine("文档工具Ribbon界面创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建文档工具Ribbon界面时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建动态工具Ribbon界面
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateDynamicToolsRibbon()
        {
            try
            {
                Console.WriteLine("开始创建动态工具Ribbon界面...");

                // 创建动态工具选项卡XML
                string ribbonXml = _ribbonManager.CreateDynamicToolsTab();

                if (string.IsNullOrEmpty(ribbonXml))
                {
                    Console.WriteLine("创建动态工具选项卡XML失败");
                    return false;
                }

                // 验证XML
                bool isValid = _ribbonManager.ValidateRibbonXml(ribbonXml);
                if (!isValid)
                {
                    Console.WriteLine("动态工具选项卡XML验证失败");
                    return false;
                }

                Console.WriteLine("动态工具Ribbon界面创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建动态工具Ribbon界面时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成完整的Ribbon界面解决方案
        /// </summary>
        /// <param name="solutionName">解决方案名称</param>
        /// <param name="customTabs">自定义选项卡定义</param>
        /// <returns>Ribbon界面解决方案</returns>
        public RibbonUISolution GenerateRibbonSolution(string solutionName, List<CustomRibbonTab> customTabs)
        {
            var solution = new RibbonUISolution
            {
                SolutionName = solutionName,
                CreatedDate = DateTime.Now
            };

            try
            {
                Console.WriteLine($"开始生成Ribbon界面解决方案: {solutionName}");

                // 为每个自定义选项卡生成XML和回调函数
                foreach (var tab in customTabs)
                {
                    // 创建选项卡XML
                    string ribbonXml = _ribbonManager.CreateCustomTabXml(
                        tab.Id,
                        tab.Label,
                        tab.Groups);

                    if (!string.IsNullOrEmpty(ribbonXml))
                    {
                        solution.RibbonXmls.Add(tab.Id, ribbonXml);
                    }

                    // 收集控件定义用于生成回调函数
                    var allControls = new List<RibbonControlDefinition>();
                    foreach (var group in tab.Groups)
                    {
                        allControls.AddRange(group.Controls);
                    }

                    // 生成回调函数模板
                    string callbackCode = _ribbonManager.GenerateCallbackTemplates(allControls);
                    if (!string.IsNullOrEmpty(callbackCode))
                    {
                        solution.CallbackCodes.Add(tab.Id, callbackCode);
                    }
                }

                // 验证所有XML
                foreach (var xml in solution.RibbonXmls.Values)
                {
                    bool isValid = _ribbonManager.ValidateRibbonXml(xml);
                    solution.ValidationResults.Add(isValid);
                }

                solution.IsComplete = solution.RibbonXmls.Any() && solution.ValidationResults.All(v => v);

                Console.WriteLine($"Ribbon界面解决方案生成完成: {solutionName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成Ribbon界面解决方案时出错: {ex.Message}");
                solution.ErrorMessage = ex.Message;
            }

            return solution;
        }

        /// <summary>
        /// 创建企业文档工具解决方案
        /// </summary>
        /// <returns>Ribbon界面解决方案</returns>
        public RibbonUISolution CreateEnterpriseDocumentSolution()
        {
            try
            {
                var customTabs = new List<CustomRibbonTab>
                {
                    new CustomRibbonTab
                    {
                        Id = "tabEnterpriseTools",
                        Label = "企业工具",
                        Groups = new List<RibbonGroupDefinition>
                        {
                            new RibbonGroupDefinition
                            {
                                Id = "grpTemplates",
                                Label = "模板",
                                Controls = new List<RibbonControlDefinition>
                                {
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnContractTemplate",
                                        Label = "合同模板",
                                        Size = "large",
                                        OnAction = "OnContractTemplate",
                                        ImageMso = "Template"
                                    },
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnReportTemplate",
                                        Label = "报告模板",
                                        Size = "large",
                                        OnAction = "OnReportTemplate",
                                        ImageMso = "Template"
                                    }
                                }
                            },
                            new RibbonGroupDefinition
                            {
                                Id = "grpReview",
                                Label = "审阅",
                                Controls = new List<RibbonControlDefinition>
                                {
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnTrackChanges",
                                        Label = "修订",
                                        OnAction = "OnTrackChanges",
                                        ImageMso = "ReviewTrackChanges"
                                    },
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnAddComment",
                                        Label = "添加批注",
                                        OnAction = "OnAddComment",
                                        ImageMso = "ReviewNewComment"
                                    }
                                }
                            }
                        }
                    }
                };

                return GenerateRibbonSolution("企业文档工具", customTabs);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建企业文档工具解决方案时出错: {ex.Message}");
                return new RibbonUISolution { ErrorMessage = ex.Message };
            }
        }

        /// <summary>
        /// 创建教育工具解决方案
        /// </summary>
        /// <returns>Ribbon界面解决方案</returns>
        public RibbonUISolution CreateEducationToolsSolution()
        {
            try
            {
                var customTabs = new List<CustomRibbonTab>
                {
                    new CustomRibbonTab
                    {
                        Id = "tabEducationTools",
                        Label = "教学工具",
                        Groups = new List<RibbonGroupDefinition>
                        {
                            new RibbonGroupDefinition
                            {
                                Id = "grpExercises",
                                Label = "练习",
                                Controls = new List<RibbonControlDefinition>
                                {
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnGrammarExercise",
                                        Label = "语法练习",
                                        Size = "large",
                                        OnAction = "OnGrammarExercise",
                                        ImageMso = "Education"
                                    },
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnVocabularyExercise",
                                        Label = "词汇练习",
                                        Size = "large",
                                        OnAction = "OnVocabularyExercise",
                                        ImageMso = "Education"
                                    }
                                }
                            },
                            new RibbonGroupDefinition
                            {
                                Id = "grpAssessment",
                                Label = "评估",
                                Controls = new List<RibbonControlDefinition>
                                {
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnCreateQuiz",
                                        Label = "创建测验",
                                        OnAction = "OnCreateQuiz",
                                        ImageMso = "Quiz"
                                    },
                                    new RibbonControlDefinition
                                    {
                                        ControlType = RibbonControlType.Button,
                                        Id = "btnGradePaper",
                                        Label = "批改试卷",
                                        OnAction = "OnGradePaper",
                                        ImageMso = "Grading"
                                    }
                                }
                            }
                        }
                    }
                };

                return GenerateRibbonSolution("教学工具", customTabs);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建教育工具解决方案时出错: {ex.Message}");
                return new RibbonUISolution { ErrorMessage = ex.Message };
            }
        }

        /// <summary>
        /// 保存Ribbon XML到文件
        /// </summary>
        /// <param name="solution">Ribbon界面解决方案</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>是否保存成功</returns>
        public bool SaveRibbonXmlFiles(RibbonUISolution solution, string outputDirectory)
        {
            try
            {
                Console.WriteLine($"开始保存Ribbon XML文件到: {outputDirectory}");

                // 确保输出目录存在
                System.IO.Directory.CreateDirectory(outputDirectory);

                // 保存每个XML文件
                foreach (var kvp in solution.RibbonXmls)
                {
                    string filePath = System.IO.Path.Combine(outputDirectory, $"{kvp.Key}.xml");
                    System.IO.File.WriteAllText(filePath, kvp.Value, Encoding.UTF8);
                    Console.WriteLine($"已保存: {filePath}");
                }

                Console.WriteLine("Ribbon XML文件保存完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存Ribbon XML文件时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 保存回调函数代码到文件
        /// </summary>
        /// <param name="solution">Ribbon界面解决方案</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>是否保存成功</returns>
        public bool SaveCallbackCodeFiles(RibbonUISolution solution, string outputDirectory)
        {
            try
            {
                Console.WriteLine($"开始保存回调函数代码文件到: {outputDirectory}");

                // 确保输出目录存在
                System.IO.Directory.CreateDirectory(outputDirectory);

                // 保存每个回调函数代码文件
                foreach (var kvp in solution.CallbackCodes)
                {
                    string filePath = System.IO.Path.Combine(outputDirectory, $"{kvp.Key}_Callbacks.cs");
                    System.IO.File.WriteAllText(filePath, kvp.Value, Encoding.UTF8);
                    Console.WriteLine($"已保存: {filePath}");
                }

                Console.WriteLine("回调函数代码文件保存完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存回调函数代码文件时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建Ribbon界面演示文档
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateRibbonDemoDocument()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                app.Visible = true;

                var document = app.ActiveDocument;

                // 创建演示文档内容
                document.Range().Text = "Ribbon界面定制演示文档\n\n" +
                                      "本文档演示了如何为Word创建自定义Ribbon界面。\n\n" +
                                      "主要内容包括：\n" +
                                      "1. 自定义选项卡和组的设计\n" +
                                      "2. 各种Ribbon控件的使用\n" +
                                      "3. 动态UI更新的实现\n" +
                                      "4. 回调函数的编写\n\n" +
                                      "要实现完整的Ribbon定制功能，需要在VSTO插件环境中进行开发。";

                // 格式化标题
                var titleRange = document.Range(0, 18);
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 格式化列表
                var listStart = document.Range().Text.IndexOf("主要内容包括：");
                var listEnd = document.Range().Text.IndexOf("要实现完整的Ribbon定制功能");
                if (listStart > 0 && listEnd > listStart)
                {
                    var listRange = document.Range(listStart, listEnd);
                    listRange.ListFormat.ApplyNumberDefault();
                }

                // 保存文档
                string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RibbonDemoDocument.docx");
                document.SaveAs(filePath);

                Console.WriteLine($"Ribbon界面演示文档已创建: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建Ribbon界面演示文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成Ribbon定制指南
        /// </summary>
        /// <returns>Ribbon定制指南</returns>
        public RibbonCustomizationGuide GenerateCustomizationGuide()
        {
            var guide = new RibbonCustomizationGuide
            {
                Title = "Ribbon界面定制完整指南",
                Version = "1.0",
                CreatedDate = DateTime.Now
            };

            try
            {
                Console.WriteLine("开始生成Ribbon定制指南...");

                // 添加章节
                guide.Sections.Add(new GuideSection
                {
                    Title = "1. Ribbon基础概念",
                    Content = "Ribbon是Office 2007及以后版本的用户界面核心组件，它将命令组织在选项卡和组中，使用户能够更轻松地找到和使用功能。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "2. Ribbon XML结构",
                    Content = "Ribbon界面通过XML定义，主要元素包括customUI、ribbon、tabs、tab、group和各种控件元素。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "3. 控件类型",
                    Content = "Ribbon支持多种控件类型，包括按钮、切换按钮、下拉列表、编辑框等，每种控件都有特定的属性和行为。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "4. 回调函数",
                    Content = "回调函数是Ribbon交互的核心，包括onAction、getEnabled、getPressed、onChange等，用于处理用户操作和动态更新界面。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "5. 动态UI更新",
                    Content = "通过getEnabled、getPressed等回调函数，可以实现Ribbon控件状态的动态更新，提供更好的用户体验。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "6. 开发环境搭建",
                    Content = "完整的Ribbon定制需要在VSTO(Visual Studio Tools for Office)环境中进行，需要安装Visual Studio和Office开发工具。"
                });

                guide.Sections.Add(new GuideSection
                {
                    Title = "7. 最佳实践",
                    Content = "设计简洁直观的界面，合理组织功能，遵循Office UI规范，优化性能，提供良好的用户体验。"
                });

                Console.WriteLine("Ribbon定制指南生成完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成Ribbon定制指南时出错: {ex.Message}");
                guide.ErrorMessage = ex.Message;
            }

            return guide;
        }
    }

    /// <summary>
    /// 自定义Ribbon选项卡类
    /// </summary>
    public class CustomRibbonTab
    {
        /// <summary>
        /// 选项卡ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 选项卡标签
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 组定义列表
        /// </summary>
        public List<RibbonGroupDefinition> Groups { get; set; } = new List<RibbonGroupDefinition>();
    }

    /// <summary>
    /// Ribbon界面解决方案类
    /// </summary>
    public class RibbonUISolution
    {
        /// <summary>
        /// 解决方案名称
        /// </summary>
        public string SolutionName { get; set; }

        /// <summary>
        /// 创建日期
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Ribbon XML字典
        /// </summary>
        public Dictionary<string, string> RibbonXmls { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 回调函数代码字典
        /// </summary>
        public Dictionary<string, string> CallbackCodes { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 验证结果列表
        /// </summary>
        public List<bool> ValidationResults { get; set; } = new List<bool>();

        /// <summary>
        /// 是否完整
        /// </summary>
        public bool IsComplete { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成解决方案报告
        /// </summary>
        /// <returns>解决方案报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"解决方案生成失败: {ErrorMessage}";
            }

            return $"Ribbon界面解决方案报告:\n" +
                   $"  解决方案名称: {SolutionName}\n" +
                   $"  创建日期: {CreatedDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  XML文件数量: {RibbonXmls.Count}\n" +
                   $"  回调代码文件数量: {CallbackCodes.Count}\n" +
                   $"  验证通过: {ValidationResults.All(v => v)}\n" +
                   $"  状态: {(IsComplete ? "完整" : "不完整")}";
        }
    }

    /// <summary>
    /// Ribbon定制指南类
    /// </summary>
    public class RibbonCustomizationGuide
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 版本
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// 创建日期
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// 章节列表
        /// </summary>
        public List<GuideSection> Sections { get; set; } = new List<GuideSection>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成定制指南报告
        /// </summary>
        /// <returns>定制指南报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"定制指南生成失败: {ErrorMessage}";
            }

            var sectionTitles = string.Join("\n  ", Sections.Select(s => s.Title));

            return $"Ribbon定制指南报告:\n" +
                   $"  标题: {Title}\n" +
                   $"  版本: {Version}\n" +
                   $"  创建日期: {CreatedDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  章节:\n  {sectionTitles}";
        }
    }

    /// <summary>
    /// 指南章节类
    /// </summary>
    public class GuideSection
    {
        /// <summary>
        /// 章节标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 章节内容
        /// </summary>
        public string Content { get; set; }
    }
}
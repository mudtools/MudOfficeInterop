using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RibbonCustomizationSample
{
    /// <summary>
    /// Ribbon定制助手类
    /// </summary>
    public class RibbonCustomizationHelper
    {
        private readonly IWordApplication _application;
        private readonly RibbonManager _ribbonManager;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        public RibbonCustomizationHelper(IWordApplication application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _ribbonManager = new RibbonManager(application);
        }

        /// <summary>
        /// 创建完整的Ribbon定制解决方案
        /// </summary>
        /// <param name="solutionName">解决方案名称</param>
        /// <param name="tabs">选项卡定义列表</param>
        /// <returns>Ribbon定制解决方案</returns>
        public RibbonCustomizationSolution CreateCustomizationSolution(string solutionName, List<RibbonTabDefinition> tabs)
        {
            var solution = new RibbonCustomizationSolution
            {
                SolutionName = solutionName,
                CreatedDate = DateTime.Now
            };

            try
            {
                Console.WriteLine($"开始创建Ribbon定制解决方案: {solutionName}");

                // 为每个选项卡生成XML
                foreach (var tab in tabs)
                {
                    var xml = _ribbonManager.CreateCustomTabXml(tab.Id, tab.Label, tab.Groups);
                    if (!string.IsNullOrEmpty(xml))
                    {
                        solution.RibbonXmlFiles.Add($"{tab.Id}.xml", xml);
                    }
                }

                // 收集所有控件定义用于生成回调函数
                var allControls = new List<RibbonControlDefinition>();
                foreach (var tab in tabs)
                {
                    foreach (var group in tab.Groups)
                    {
                        allControls.AddRange(group.Controls);
                    }
                }

                // 生成回调函数模板
                var callbackCode = _ribbonManager.GenerateCallbackTemplates(allControls);
                if (!string.IsNullOrEmpty(callbackCode))
                {
                    solution.CallbackCode = callbackCode;
                }

                // 验证生成的XML
                foreach (var xml in solution.RibbonXmlFiles.Values)
                {
                    bool isValid = _ribbonManager.ValidateRibbonXml(xml);
                    solution.ValidationResults.Add(isValid);
                }

                solution.IsComplete = solution.RibbonXmlFiles.Any() && solution.ValidationResults.All(v => v);
                
                Console.WriteLine($"Ribbon定制解决方案创建完成: {solutionName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建Ribbon定制解决方案时出错: {ex.Message}");
                solution.ErrorMessage = ex.Message;
            }

            return solution;
        }

        /// <summary>
        /// 创建文档处理工具解决方案
        /// </summary>
        /// <returns>Ribbon定制解决方案</returns>
        public RibbonCustomizationSolution CreateDocumentToolsSolution()
        {
            try
            {
                var tabs = new List<RibbonTabDefinition>
                {
                    new RibbonTabDefinition
                    {
                        Id = "tabDocumentTools",
                        Label = "文档工具",
                        Groups = new List<RibbonGroupDefinition>
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
                                    }
                                }
                            }
                        }
                    }
                };

                return CreateCustomizationSolution("文档处理工具", tabs);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建文档处理工具解决方案时出错: {ex.Message}");
                return new RibbonCustomizationSolution { ErrorMessage = ex.Message };
            }
        }

        /// <summary>
        /// 创建动态工具解决方案
        /// </summary>
        /// <returns>Ribbon定制解决方案</returns>
        public RibbonCustomizationSolution CreateDynamicToolsSolution()
        {
            try
            {
                var tabs = new List<RibbonTabDefinition>
                {
                    new RibbonTabDefinition
                    {
                        Id = "tabDynamicTools",
                        Label = "动态工具",
                        Groups = new List<RibbonGroupDefinition>
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
                                    }
                                }
                            }
                        }
                    }
                };

                return CreateCustomizationSolution("动态工具", tabs);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建动态工具解决方案时出错: {ex.Message}");
                return new RibbonCustomizationSolution { ErrorMessage = ex.Message };
            }
        }

        /// <summary>
        /// 生成完整的VSTO插件项目结构
        /// </summary>
        /// <param name="solution">Ribbon定制解决方案</param>
        /// <returns>VSTO项目结构</returns>
        public VstoProjectStructure GenerateVstoProjectStructure(RibbonCustomizationSolution solution)
        {
            var project = new VstoProjectStructure
            {
                ProjectName = solution.SolutionName,
                CreatedDate = solution.CreatedDate
            };

            try
            {
                Console.WriteLine($"开始生成VSTO项目结构: {solution.SolutionName}");

                // 添加项目文件
                project.ProjectFiles.Add("ThisAddIn.cs", GenerateThisAddInCode());
                project.ProjectFiles.Add("Ribbon.cs", GenerateRibbonCode(solution));
                project.ProjectFiles.Add($"{solution.SolutionName}Ribbon.xml", GetFirstRibbonXml(solution));

                // 添加项目配置文件
                project.ProjectFiles.Add($"{solution.SolutionName}.csproj", GenerateProjectFile(solution));

                Console.WriteLine($"VSTO项目结构生成完成: {solution.SolutionName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成VSTO项目结构时出错: {ex.Message}");
                project.ErrorMessage = ex.Message;
            }

            return project;
        }

        /// <summary>
        /// 生成ThisAddIn.cs代码
        /// </summary>
        /// <returns>ThisAddIn.cs代码</returns>
        private string GenerateThisAddInCode()
        {
            return @"using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace [SolutionName]
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}";
        }

        /// <summary>
        /// 生成Ribbon.cs代码
        /// </summary>
        /// <param name="solution">Ribbon定制解决方案</param>
        /// <returns>Ribbon.cs代码</returns>
        private string GenerateRibbonCode(RibbonCustomizationSolution solution)
        {
            var codeBuilder = new StringBuilder();
            codeBuilder.AppendLine(@"using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace [SolutionName]
{
    [ComVisible(true)]
    public class Ribbon : OfficeRibbon
    {
        private Microsoft.Office.Core.IRibbonUI ribbonUI;

        public Ribbon()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.tabDocumentTools = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.grpFormatting = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnHeading1 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.btnHeading2 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.grpTables = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnInsertTable = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tabDocumentTools.SuspendLayout();
            this.grpFormatting.SuspendLayout();
            this.grpTables.SuspendLayout();
            this.SuspendLayout();
            
            // 
            // tabDocumentTools
            // 
            this.tabDocumentTools.Groups.Add(this.grpFormatting);
            this.tabDocumentTools.Groups.Add(this.grpTables);
            this.tabDocumentTools.Label = ""文档工具"";
            this.tabDocumentTools.Name = ""tabDocumentTools"";
            // 
            // grpFormatting
            // 
            this.grpFormatting.Items.Add(this.btnHeading1);
            this.grpFormatting.Items.Add(this.btnHeading2);
            this.grpFormatting.Label = ""格式化工具"";
            this.grpFormatting.Name = ""grpFormatting"";
            // 
            // btnHeading1
            // 
            this.btnHeading1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHeading1.Image = global::[SolutionName].Properties.Resources.StyleHeading1;
            this.btnHeading1.Label = ""标题1"";
            this.btnHeading1.Name = ""btnHeading1"";
            this.btnHeading1.ShowImage = true;
            this.btnHeading1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnHeading1);
            // 
            // btnHeading2
            // 
            this.btnHeading2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHeading2.Image = global::[SolutionName].Properties.Resources.StyleHeading2;
            this.btnHeading2.Label = ""标题2"";
            this.btnHeading2.Name = ""btnHeading2"";
            this.btnHeading2.ShowImage = true;
            this.btnHeading2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnHeading2);
            // 
            // grpTables
            // 
            this.grpTables.Items.Add(this.btnInsertTable);
            this.grpTables.Label = ""表格工具"";
            this.grpTables.Name = ""grpTables"";
            // 
            // btnInsertTable
            // 
            this.btnInsertTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertTable.Image = global::[SolutionName].Properties.Resources.TableInsertTable;
            this.btnInsertTable.Label = ""插入表格"";
            this.btnInsertTable.Name = ""btnInsertTable"";
            this.btnInsertTable.ShowImage = true;
            this.btnInsertTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnInsertTable);
            // 
            // Ribbon
            // 
            this.Name = ""Ribbon"";
            this.RibbonType = ""Microsoft.Word.Document"";
            this.Tabs.Add(this.tabDocumentTools);
            this.tabDocumentTools.ResumeLayout(false);
            this.tabDocumentTools.PerformLayout();
            this.grpFormatting.ResumeLayout(false);
            this.grpFormatting.PerformLayout();
            this.grpTables.ResumeLayout(false);
            this.grpTables.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDocumentTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHeading1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHeading2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertTable;

        public void OnRibbonLoad(Microsoft.Office.Core.IRibbonUI ribbon)
        {
            this.ribbonUI = ribbon;
        }

        public void RefreshRibbon()
        {
            ribbonUI?.Invalidate();
        }
");
            
            // 添加解决方案中的回调函数
            if (!string.IsNullOrEmpty(solution.CallbackCode))
            {
                codeBuilder.AppendLine(solution.CallbackCode);
            }
            else
            {
                codeBuilder.AppendLine(@"
        public void OnHeading1(object sender, RibbonControlEventArgs e)
        {
            // 实现标题1功能
            MessageBox.Show(""应用标题1样式"");
        }

        public void OnHeading2(object sender, RibbonControlEventArgs e)
        {
            // 实现标题2功能
            MessageBox.Show(""应用标题2样式"");
        }

        public void OnInsertTable(object sender, RibbonControlEventArgs e)
        {
            // 实现插入表格功能
            MessageBox.Show(""插入表格"");
        }
");
            }

            codeBuilder.AppendLine(@"    }
}");
            return codeBuilder.ToString();
        }

        /// <summary>
        /// 获取第一个Ribbon XML
        /// </summary>
        /// <param name="solution">Ribbon定制解决方案</param>
        /// <returns>Ribbon XML</returns>
        private string GetFirstRibbonXml(RibbonCustomizationSolution solution)
        {
            return solution.RibbonXmlFiles.Values.FirstOrDefault() ?? string.Empty;
        }

        /// <summary>
        /// 生成项目文件
        /// </summary>
        /// <param name="solution">Ribbon定制解决方案</param>
        /// <returns>项目文件内容</returns>
        private string GenerateProjectFile(RibbonCustomizationSolution solution)
        {
            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Project ToolsVersion=""15.0"" xmlns=""http://schemas.microsoft.com/developer/msbuild/2003"">
  <PropertyGroup>
    <ProjectTypeGuids>{{BAA0C2D2-18E2-41B9-852F-F413020CAA33}};{{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}}</ProjectTypeGuids>
    <Configuration Condition="" '$(Configuration)' == '' "">Debug</Configuration>
    <Platform Condition="" '$(Platform)' == '' "">AnyCPU</Platform>
    <OutputType>Library</OutputType>
    <RootNamespace>{solution.SolutionName}</RootNamespace>
    <AssemblyName>{solution.SolutionName}</AssemblyName>
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <ProjectExtensions />
  </PropertyGroup>
  <ItemGroup>
    <Reference Include=""System"" />
    <Reference Include=""System.Data"" />
    <Reference Include=""System.Drawing"" />
    <Reference Include=""System.Windows.Forms"" />
    <Reference Include=""System.Xml"" />
    <Reference Include=""Microsoft.Office.Tools.Common.v4.0.Utilities, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"">
      <Private>True</Private>
    </Reference>
  </ItemGroup>
  <ItemGroup>
    <Reference Include=""Microsoft.Office.Tools.Word, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"" />
    <Reference Include=""Microsoft.Office.Tools.Word.v4.0.Utilities, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"">
      <Private>True</Private>
    </Reference>
    <Reference Include=""Microsoft.Office.Tools.Common, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"" />
    <Reference Include=""Microsoft.VisualStudio.Tools.Applications.Runtime, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"" />
    <Reference Include=""Office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"">
      <Private>False</Private>
    </Reference>
    <Reference Include=""Microsoft.Office.Interop.Word, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"">
      <Private>False</Private>
    </Reference>
    <Reference Include=""stdole, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"">
      <Private>False</Private>
    </Reference>
  </ItemGroup>
  <ItemGroup>
    <None Include=""ThisAddIn.Designer.xml"">
      <DependentUpon>ThisAddIn.cs</DependentUpon>
    </None>
    <Compile Include=""ThisAddIn.cs"">
      <SubType>Code</SubType>
    </Compile>
    <None Include=""[solution.SolutionName]Ribbon.xml"">
      <CopyToOutputDirectory>Always</CopyToOutputDirectory>
    </None>
    <Compile Include=""Ribbon.cs">
      <SubType>Component</SubType>
    </Compile>
  </ItemGroup>
</Project>";
        }

        /// <summary>
        /// 评估Ribbon定制复杂度
        /// </summary>
        /// <param name="solution">Ribbon定制解决方案</param>
        /// <returns>复杂度评估结果</returns>
        public RibbonComplexityAssessment AssessComplexity(RibbonCustomizationSolution solution)
        {
            var assessment = new RibbonComplexityAssessment
            {
                SolutionName = solution.SolutionName
            };

            try
            {
                // 计算选项卡数量
                assessment.TabCount = solution.RibbonXmlFiles.Count;

                // 计算控件总数
                assessment.ControlCount = solution.RibbonXmlFiles.Values
                    .Sum(xml => xml.Split(new[] { "<button", "<toggleButton", "<dropDown", "<editBox" }, StringSplitOptions.None).Length - 1);

                // 计算回调函数数量
                assessment.CallbackCount = solution.CallbackCode?.Split(new[] { "public void", "public bool" }, StringSplitOptions.None).Length - 1 ?? 0;

                // 评估复杂度等级
                if (assessment.ControlCount <= 5 && assessment.CallbackCount <= 5)
                {
                    assessment.ComplexityLevel = "简单";
                }
                else if (assessment.ControlCount <= 15 && assessment.CallbackCount <= 15)
                {
                    assessment.ComplexityLevel = "中等";
                }
                else
                {
                    assessment.ComplexityLevel = "复杂";
                }

                Console.WriteLine("Ribbon定制复杂度评估完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"评估Ribbon定制复杂度时出错: {ex.Message}");
                assessment.ErrorMessage = ex.Message;
            }

            return assessment;
        }
    }

    /// <summary>
    /// Ribbon选项卡定义类
    /// </summary>
    public class RibbonTabDefinition
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
    /// Ribbon定制解决方案类
    /// </summary>
    public class RibbonCustomizationSolution
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
        /// Ribbon XML文件字典
        /// </summary>
        public Dictionary<string, string> RibbonXmlFiles { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 回调函数代码
        /// </summary>
        public string CallbackCode { get; set; }

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

            return $"Ribbon定制解决方案报告:\n" +
                   $"  解决方案名称: {SolutionName}\n" +
                   $"  创建日期: {CreatedDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  XML文件数量: {RibbonXmlFiles.Count}\n" +
                   $"  回调函数代码行数: {CallbackCode?.Split('\n').Length ?? 0}\n" +
                   $"  验证通过: {ValidationResults.All(v => v)}\n" +
                   $"  状态: {(IsComplete ? "完整" : "不完整")}";
        }
    }

    /// <summary>
    /// VSTO项目结构类
    /// </summary>
    public class VstoProjectStructure
    {
        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; }

        /// <summary>
        /// 创建日期
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// 项目文件字典
        /// </summary>
        public Dictionary<string, string> ProjectFiles { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成项目结构报告
        /// </summary>
        /// <returns>项目结构报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"项目结构生成失败: {ErrorMessage}";
            }

            var fileNames = string.Join(", ", ProjectFiles.Keys);

            return $"VSTO项目结构报告:\n" +
                   $"  项目名称: {ProjectName}\n" +
                   $"  创建日期: {CreatedDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  项目文件: {fileNames}\n" +
                   $"  文件数量: {ProjectFiles.Count}";
        }
    }

    /// <summary>
    /// Ribbon复杂度评估类
    /// </summary>
    public class RibbonComplexityAssessment
    {
        /// <summary>
        /// 解决方案名称
        /// </summary>
        public string SolutionName { get; set; }

        /// <summary>
        /// 选项卡数量
        /// </summary>
        public int TabCount { get; set; }

        /// <summary>
        /// 控件数量
        /// </summary>
        public int ControlCount { get; set; }

        /// <summary>
        /// 回调函数数量
        /// </summary>
        public int CallbackCount { get; set; }

        /// <summary>
        /// 复杂度等级
        /// </summary>
        public string ComplexityLevel { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成复杂度评估报告
        /// </summary>
        /// <returns>复杂度评估报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"复杂度评估失败: {ErrorMessage}";
            }

            return $"Ribbon复杂度评估报告:\n" +
                   $"  解决方案名称: {SolutionName}\n" +
                   $"  选项卡数量: {TabCount}\n" +
                   $"  控件数量: {ControlCount}\n" +
                   $"  回调函数数量: {CallbackCount}\n" +
                   $"  复杂度等级: {ComplexityLevel}";
        }
    }
}
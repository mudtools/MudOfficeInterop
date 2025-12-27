using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskPanesAndDialogsSample
{
    /// <summary>
    /// 任务窗格管理器类
    /// </summary>
    /// <remarks>
    /// 注意：此示例展示了任务窗格的概念和使用方法。
    /// 实际的任务窗格功能需要在VSTO插件环境中实现。
    /// </remarks>
    public class TaskPaneManager
    {
        private readonly IWordApplication _application;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        public TaskPaneManager(IWordApplication? Application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// 获取任务窗格信息
        /// </summary>
        /// <returns>任务窗格信息</returns>
        public TaskPaneInfo GetTaskPaneInfo()
        {
            var info = new TaskPaneInfo();

            try
            {
                // 获取任务窗格集合信息
                // 注意：在纯COM互操作环境中，任务窗格功能有限
                // 完整的任务窗格功能需要在VSTO插件环境中实现
                info.TaskPaneCount = 0; // 在纯COM环境中无法直接访问任务窗格
                info.IsTaskPaneSupported = false;
                info.Message = "任务窗格功能需要在VSTO插件环境中实现";

                Console.WriteLine($"任务窗格数量: {info.TaskPaneCount}");
                Console.WriteLine($"是否支持任务窗格: {info.IsTaskPaneSupported}");
                Console.WriteLine(info.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取任务窗格信息时出错: {ex.Message}");
                info.ErrorMessage = ex.Message;
            }

            return info;
        }

        /// <summary>
        /// 创建任务窗格XML定义
        /// </summary>
        /// <param name="tabId">选项卡ID</param>
        /// <param name="tabLabel">选项卡标签</param>
        /// <param name="groupId">组ID</param>
        /// <param name="groupLabel">组标签</param>
        /// <param name="buttonId">按钮ID</param>
        /// <param name="buttonLabel">按钮标签</param>
        /// <returns>任务窗格XML定义</returns>
        public string CreateTaskPaneXml(
            string tabId,
            string tabLabel,
            string groupId,
            string groupLabel,
            string buttonId,
            string buttonLabel)
        {
            var xmlBuilder = new StringBuilder();
            xmlBuilder.AppendLine("<?xml version='1.0' encoding='utf-8' ?>");
            xmlBuilder.AppendLine("<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>");
            xmlBuilder.AppendLine("  <ribbon>");
            xmlBuilder.AppendLine("    <tabs>");
            xmlBuilder.AppendLine($"      <tab id='{tabId}' label='{tabLabel}'>");
            xmlBuilder.AppendLine($"        <group id='{groupId}' label='{groupLabel}'>");
            xmlBuilder.AppendLine($"          <button id='{buttonId}' label='{buttonLabel}' onAction='OnShowTaskPane' />");
            xmlBuilder.AppendLine("        </group>");
            xmlBuilder.AppendLine("      </tab>");
            xmlBuilder.AppendLine("    </tabs>");
            xmlBuilder.AppendLine("  </ribbon>");
            xmlBuilder.AppendLine("</customUI>");

            Console.WriteLine("已生成任务窗格XML定义:");
            Console.WriteLine(xmlBuilder.ToString());

            return xmlBuilder.ToString();
        }

        /// <summary>
        /// 生成任务窗格用户控件代码
        /// </summary>
        /// <returns>用户控件代码</returns>
        public string GenerateTaskPaneUserControlCode()
        {
            var codeBuilder = new StringBuilder();
            codeBuilder.AppendLine("任务窗格用户控件示例代码 (C# Windows Forms):");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("public partial class DocumentToolsPane : UserControl");
            codeBuilder.AppendLine("{");
            codeBuilder.AppendLine("    public DocumentToolsPane()");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        InitializeComponent();");
            codeBuilder.AppendLine("        InitializeCustomComponents();");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void InitializeCustomComponents()");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        // 创建控件");
            codeBuilder.AppendLine("        var groupBox = new GroupBox();");
            codeBuilder.AppendLine("        groupBox.Text = \"文档格式化\";");
            codeBuilder.AppendLine("        groupBox.Location = new Point(10, 10);");
            codeBuilder.AppendLine("        groupBox.Size = new Size(280, 150);");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        var lblFontSize = new Label();");
            codeBuilder.AppendLine("        lblFontSize.Text = \"字体大小:\";");
            codeBuilder.AppendLine("        lblFontSize.Location = new Point(10, 20);");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        var cmbFontSize = new ComboBox();");
            codeBuilder.AppendLine("        cmbFontSize.Items.AddRange(new[] { \"8\", \"10\", \"12\", \"14\", \"16\", \"18\", \"20\" });");
            codeBuilder.AppendLine("        cmbFontSize.Location = new Point(80, 20);");
            codeBuilder.AppendLine("        cmbFontSize.SelectedIndex = 2; // 默认选择12");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        var btnBold = new Button();");
            codeBuilder.AppendLine("        btnBold.Text = \"加粗\";");
            codeBuilder.AppendLine("        btnBold.Location = new Point(10, 50);");
            codeBuilder.AppendLine("        btnBold.Click += BtnBold_Click;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        var btnItalic = new Button();");
            codeBuilder.AppendLine("        btnItalic.Text = \"斜体\";");
            codeBuilder.AppendLine("        btnItalic.Location = new Point(90, 50);");
            codeBuilder.AppendLine("        btnItalic.Click += BtnItalic_Click;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        var btnUnderline = new Button();");
            codeBuilder.AppendLine("        btnUnderline.Text = \"下划线\";");
            codeBuilder.AppendLine("        btnUnderline.Location = new Point(170, 50);");
            codeBuilder.AppendLine("        btnUnderline.Click += BtnUnderline_Click;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        // 添加控件到GroupBox");
            codeBuilder.AppendLine("        groupBox.Controls.AddRange(new Control[] { lblFontSize, cmbFontSize, btnBold, btnItalic, btnUnderline });");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        // 添加到用户控件");
            codeBuilder.AppendLine("        this.Controls.Add(groupBox);");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void BtnBold_Click(object sender, EventArgs e)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        ApplyFormatting(\"Bold\");");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void BtnItalic_Click(object sender, EventArgs e)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        ApplyFormatting(\"Italic\");");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void BtnUnderline_Click(object sender, EventArgs e)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        ApplyFormatting(\"Underline\");");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void ApplyFormatting(string formatType)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        try");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            // 在VSTO插件中，可以通过Globals.ThisAddIn.Application访问Word应用程序");
            codeBuilder.AppendLine("            // var app = Globals.ThisAddIn.Application;");
            codeBuilder.AppendLine("            // var selection = app.Selection;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("            // 示例代码:");
            codeBuilder.AppendLine("            // if (selection != null)");
            codeBuilder.AppendLine("            // {");
            codeBuilder.AppendLine("            //     switch (formatType)");
            codeBuilder.AppendLine("            //     {");
            codeBuilder.AppendLine("            //         case \"Bold\":");
            codeBuilder.AppendLine("            //             selection.Font.Bold = (selection.Font.Bold == 1) ? 0 : 1;");
            codeBuilder.AppendLine("            //             break;");
            codeBuilder.AppendLine("            //         case \"Italic\":");
            codeBuilder.AppendLine("            //             selection.Font.Italic = (selection.Font.Italic == 1) ? 0 : 1;");
            codeBuilder.AppendLine("            //             break;");
            codeBuilder.AppendLine("            //         case \"Underline\":");
            codeBuilder.AppendLine("            //             selection.Font.Underline = (selection.Font.Underline == WdUnderline.wdUnderlineSingle)");
            codeBuilder.AppendLine("            //                 ? WdUnderline.wdUnderlineNone");
            codeBuilder.AppendLine("            //                 : WdUnderline.wdUnderlineSingle;");
            codeBuilder.AppendLine("            //             break;");
            codeBuilder.AppendLine("            //     }");
            codeBuilder.AppendLine("            // }");
            codeBuilder.AppendLine("        }");
            codeBuilder.AppendLine("        catch (Exception ex)");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            // MessageBox.Show($\"格式化出错: {ex.Message}\");");
            codeBuilder.AppendLine("            Console.WriteLine($\"格式化出错: {ex.Message}\");");
            codeBuilder.AppendLine("        }");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("}");

            Console.WriteLine("已生成任务窗格用户控件代码示例");
            return codeBuilder.ToString();
        }

        /// <summary>
        /// 生成VSTO插件任务窗格实现代码
        /// </summary>
        /// <returns>VSTO插件代码</returns>
        public string GenerateVstoTaskPaneImplementation()
        {
            var codeBuilder = new StringBuilder();
            codeBuilder.AppendLine("VSTO插件中任务窗格实现示例:");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("// 在VSTO插件的ThisAddIn.cs文件中添加以下代码:");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("public partial class ThisAddIn");
            codeBuilder.AppendLine("{");
            codeBuilder.AppendLine("    private Microsoft.Office.Tools.CustomTaskPane customTaskPane;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void ThisAddIn_Startup(object sender, System.EventArgs e)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        // 创建用户控件实例");
            codeBuilder.AppendLine("        DocumentToolsPane taskPaneControl = new DocumentToolsPane();");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("        // 添加自定义任务窗格");
            codeBuilder.AppendLine("        customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, \"文档工具\");");
            codeBuilder.AppendLine("        customTaskPane.Visible = true;");
            codeBuilder.AppendLine("        customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;");
            codeBuilder.AppendLine("        customTaskPane.Width = 300;");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        // 清理资源");
            codeBuilder.AppendLine("        if (customTaskPane != null)");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            customTaskPane.Dispose();");
            codeBuilder.AppendLine("        }");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    // 显示任务窗格的方法");
            codeBuilder.AppendLine("    public void ShowTaskPane()");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        if (customTaskPane != null)");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            customTaskPane.Visible = true;");
            codeBuilder.AppendLine("        }");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("    // 隐藏任务窗格的方法");
            codeBuilder.AppendLine("    public void HideTaskPane()");
            codeBuilder.AppendLine("    {");
            codeBuilder.AppendLine("        if (customTaskPane != null)");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            customTaskPane.Visible = false;");
            codeBuilder.AppendLine("        }");
            codeBuilder.AppendLine("    }");
            codeBuilder.AppendLine("}");

            Console.WriteLine("已生成VSTO插件任务窗格实现代码示例");
            return codeBuilder.ToString();
        }

        /// <summary>
        /// 生成任务窗格最佳实践指南
        /// </summary>
        /// <returns>最佳实践指南</returns>
        public string GenerateTaskPaneBestPractices()
        {
            var guideBuilder = new StringBuilder();
            guideBuilder.AppendLine("=== 任务窗格最佳实践指南 ===");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("1. 任务窗格设计原则:");
            guideBuilder.AppendLine("   - 保持界面简洁，避免过于复杂");
            guideBuilder.AppendLine("   - 提供清晰的标签和说明");
            guideBuilder.AppendLine("   - 合理组织控件布局");
            guideBuilder.AppendLine("   - 支持键盘导航");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("2. 用户体验优化:");
            guideBuilder.AppendLine("   - 提供有意义的默认值");
            guideBuilder.AppendLine("   - 给出清晰的操作反馈");
            guideBuilder.AppendLine("   - 处理异常情况");
            guideBuilder.AppendLine("   - 支持撤销操作");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("3. 性能考虑:");
            guideBuilder.AppendLine("   - 避免在UI线程执行耗时操作");
            guideBuilder.AppendLine("   - 及时释放资源");
            guideBuilder.AppendLine("   - 优化控件响应速度");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("4. 兼容性考虑:");
            guideBuilder.AppendLine("   - 支持不同的Office版本");
            guideBuilder.AppendLine("   - 适配不同的屏幕分辨率");
            guideBuilder.AppendLine("   - 考虑多显示器环境");

            Console.WriteLine("已生成任务窗格最佳实践指南");
            return guideBuilder.ToString();
        }
    }

    /// <summary>
    /// 任务窗格信息类
    /// </summary>
    public class TaskPaneInfo
    {
        /// <summary>
        /// 任务窗格数量
        /// </summary>
        public int TaskPaneCount { get; set; }

        /// <summary>
        /// 是否支持任务窗格
        /// </summary>
        public bool IsTaskPaneSupported { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}
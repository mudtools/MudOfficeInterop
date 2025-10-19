//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace TableOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 表格操作示例");

            // 示例1: IWordTable和IWordTables接口
            Console.WriteLine("\n=== 示例1: IWordTable和IWordTables接口 ===");
            WordTableInterfaceDemo();

            // 示例2: 创建和删除表格
            Console.WriteLine("\n=== 示例2: 创建和删除表格 ===");
            CreateTableDemo();

            // 示例3: 表格格式化
            Console.WriteLine("\n=== 示例3: 表格格式化 ===");
            TableFormattingDemo();

            // 示例4: 单元格操作
            Console.WriteLine("\n=== 示例4: 单元格操作 ===");
            CellOperationsDemo();

            // 示例5: 表格数据处理
            Console.WriteLine("\n=== 示例5: 表格数据处理 ===");
            TableDataProcessingDemo();

            // 示例6: 高级表格操作
            Console.WriteLine("\n=== 示例6: 高级表格操作 ===");
            AdvancedTableOperationsDemo();

            // 示例7: 实际应用示例
            Console.WriteLine("\n=== 示例7: 实际应用示例 ===");
            RealWorldTableDemo();

            // 示例8: 表格样式和主题
            Console.WriteLine("\n=== 示例8: 表格样式和主题 ===");
            TableStylingDemo();

            // 示例9: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例9: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// IWordTable和IWordTables接口示例
        /// </summary>
        static void WordTableInterfaceDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 获取表格集合
                using var tables = document.Tables;

                // 获取表格数量
                int tableCount = tables.Count;
                Console.WriteLine($"初始表格数量: {tableCount}");

                // 创建一个表格用于演示
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 2, 2);

                // 再次获取表格数量
                tableCount = tables.Count;
                Console.WriteLine($"创建表格后数量: {tableCount}");

                // 访问特定表格（索引从1开始）
                if (tableCount > 0)
                {
                    using var firstTable = tables[1];
                    Console.WriteLine("成功访问第一个表格");
                }

                Console.WriteLine("IWordTable和IWordTables接口操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"IWordTable和IWordTables接口操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建和删除表格示例
        /// </summary>
        static void CreateTableDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 方法1：在文档末尾添加表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table1 = document.Tables.Add(range, 3, 4); // 3行4列
                Console.WriteLine("在文档末尾创建了3行4列的表格");

                // 方法2：在指定位置添加表格
                using var range2 = document.Range(0, 0);
                using var table2 = document.Tables.Add(range2, 2, 3); // 2行3列
                Console.WriteLine("在文档开头创建了2行3列的表格");

                // 设置表格标题
                table1.Title = "示例表格";
                table1.Descr = "这是一个示例表格";
                Console.WriteLine("设置表格标题和描述完成");

                Console.WriteLine("创建表格操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建表格操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 表格格式化示例
        /// </summary>
        static void TableFormattingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 创建表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 4, 3);

                // 设置表格边框
                table.Borders.Enable = true;
                table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleDot;
                table.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleDot;

                // 设置表格对齐方式
                table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                // 设置表格宽度
                table.AllowAutoFit = false;
                table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;

                // 设置列宽
                table.Columns[1].Width = 100;
                table.Columns[2].Width = 150;
                table.Columns[3].Width = 200;

                Console.WriteLine("表格格式化操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"表格格式化操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 单元格操作示例
        /// </summary>
        static void CellOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 创建表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 3, 3);

                // 访问单元格
                using var cell = table.Cell(1, 1); // 第一行第一列（索引从1开始）
                using var cellRange = cell.Range;

                // 设置单元格文本
                cellRange.Text = "单元格内容";

                // 设置单元格格式
                cellRange.Font.Bold = true;
                cellRange.Font.Size = 12;
                cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 设置单元格底纹
                cell.Shading.BackgroundPatternColor = WdColor.wdColorLightBlue;

                Console.WriteLine("单元格操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"单元格操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 表格数据处理示例
        /// </summary>
        static void TableDataProcessingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 创建表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 4, 3);

                // 填充表头
                table.Cell(1, 1).Range.Text = "姓名";
                table.Cell(1, 2).Range.Text = "年龄";
                table.Cell(1, 3).Range.Text = "职业";

                // 填充数据
                string[,] data = {
                    {"张三", "25", "工程师"},
                    {"李四", "30", "设计师"},
                    {"王五", "28", "产品经理"}
                };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = data[i, j];
                    }
                }

                // 格式化表头
                for (int i = 1; i <= 3; i++)
                {
                    using var headerCell = table.Cell(1, i);
                    headerCell.Range.Font.Bold = true;
                    headerCell.Range.Font.Color = WdColor.wdColorWhite;
                    headerCell.Shading.BackgroundPatternColor = WdColor.wdColorDarkBlue;
                    headerCell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // 格式化数据行
                for (int row = 2; row <= 4; row++)
                {
                    for (int col = 1; col <= 3; col++)
                    {
                        using var cell = table.Cell(row, col);
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }

                Console.WriteLine("表格数据处理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"表格数据处理出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 高级表格操作示例
        /// </summary>
        static void AdvancedTableOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 创建带数据的表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 5, 3);

                // 填充数据
                table.Cell(1, 1).Range.Text = "产品";
                table.Cell(1, 2).Range.Text = "销量";
                table.Cell(1, 3).Range.Text = "价格";

                table.Cell(2, 1).Range.Text = "产品A";
                table.Cell(2, 2).Range.Text = "100";
                table.Cell(2, 3).Range.Text = "50";

                table.Cell(3, 1).Range.Text = "产品B";
                table.Cell(3, 2).Range.Text = "200";
                table.Cell(3, 3).Range.Text = "30";

                table.Cell(4, 1).Range.Text = "产品C";
                table.Cell(4, 2).Range.Text = "150";
                table.Cell(4, 3).Range.Text = "40";

                // 格式化表头
                for (int i = 1; i <= 3; i++)
                {
                    using var cell = table.Cell(1, i);
                    cell.Range.Font.Bold = true;
                    cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                }

                // 添加总计行
                table.Cell(5, 1).Range.Text = "总计";
                table.Cell(5, 2).Range.Text = "=SUM(ABOVE)"; // 使用公式计算总销量
                table.Cell(5, 3).Range.Text = "平均价格";

                // 更新表格中的字段（公式）
                table.Range.Fields.Update();

                Console.WriteLine("高级表格操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级表格操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldTableDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 在实际应用示例中隐藏Word窗口

                using var document = app.ActiveDocument;

                // 添加标题
                using var title = document.Range();
                title.Text = "销售数据报表\n";
                title.Font.Name = "微软雅黑";
                title.Font.Size = 18;
                title.Font.Bold = true;
                title.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title.ParagraphFormat.SpaceAfter = 24;

                // 添加报表说明
                using var description = document.Range(document.Content.End - 1, document.Content.End - 1);
                description.Text = "本报表展示了2025年各季度销售数据\n\n";
                description.Font.Name = "宋体";
                description.Font.Size = 12;
                description.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 创建销售数据表格
                using var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(tableRange, 6, 5);

                // 设置表格标题
                table.Title = "季度销售数据";
                table.Descr = "2025年各季度销售数据表";

                // 填充表头
                string[] headers = { "季度", "产品A", "产品B", "产品C", "总计" };
                for (int i = 0; i < headers.Length; i++)
                {
                    table.Cell(1, i + 1).Range.Text = headers[i];
                }

                // 填充数据
                string[,] salesData = {
                    {"Q1", "1000", "1500", "2000", "4500"},
                    {"Q2", "1200", "1800", "2200", "5200"},
                    {"Q3", "1100", "1600", "2100", "4800"},
                    {"Q4", "1300", "1900", "2300", "5500"}
                };

                for (int i = 0; i < salesData.GetLength(0); i++)
                {
                    for (int j = 0; j < salesData.GetLength(1); j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = salesData[i, j];
                    }
                }

                // 格式化表格
                // 表格边框
                table.Borders.Enable = true;
                table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;

                // 表头格式
                for (int i = 1; i <= 5; i++)
                {
                    using var cell = table.Cell(1, i);
                    cell.Range.Font.Bold = true;
                    cell.Range.Font.Color = WdColor.wdColorWhite;
                    cell.Shading.BackgroundPatternColor = WdColor.wdColorDarkBlue;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // 数据行格式
                for (int row = 2; row <= 6; row++)
                {
                    for (int col = 1; col <= 5; col++)
                    {
                        using var cell = table.Cell(row, col);
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        // 交替行颜色
                        if (row % 2 == 0)
                        {
                            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                        }
                    }
                }

                // 设置列宽
                table.AllowAutoFit = false;
                table.Columns[1].Width = 80;   // 季度列
                table.Columns[2].Width = 80;   // 产品A列
                table.Columns[3].Width = 80;   // 产品B列
                table.Columns[4].Width = 80;   // 产品C列
                table.Columns[5].Width = 80;   // 总计列

                // 更新公式字段
                table.Range.Fields.Update();

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "TableReportDemo.docx");
                document.SaveAs(filePath);

                Console.WriteLine($"表格报表已创建: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建报表时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 表格样式和主题示例
        /// </summary>
        static void TableStylingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 创建表格
                using var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                using var table = document.Tables.Add(range, 4, 3);

                // 填充数据
                table.Cell(1, 1).Range.Text = "项目";
                table.Cell(1, 2).Range.Text = "值1";
                table.Cell(1, 3).Range.Text = "值2";

                table.Cell(2, 1).Range.Text = "项目A";
                table.Cell(2, 2).Range.Text = "100";
                table.Cell(2, 3).Range.Text = "200";

                table.Cell(3, 1).Range.Text = "项目B";
                table.Cell(3, 2).Range.Text = "150";
                table.Cell(3, 3).Range.Text = "250";

                // 应用内置表格样式
                table.Style = "网格型";

                Console.WriteLine("表格样式和主题操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"表格样式和主题操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                Console.WriteLine("使用TableOperationsManager辅助类进行完整操作:");

                // 创建表格操作管理器实例
                var tableManager = new TableOperationsManager();

                // 创建简单表格
                var simpleTableResult = tableManager.CreateSimpleTable();
                Console.WriteLine($"简单表格创建结果:");
                Console.WriteLine($"  表格行数: {simpleTableResult.RowCount}");
                Console.WriteLine($"  表格列数: {simpleTableResult.ColumnCount}");

                // 创建数据表格
                var dataTableResult = tableManager.CreateDataTable();
                Console.WriteLine($"数据表格创建结果:");
                Console.WriteLine($"  表格标题: {dataTableResult.Title}");
                Console.WriteLine($"  数据行数: {dataTableResult.DataRowCount}");

                // 创建格式化表格
                var formattedTableResult = tableManager.CreateFormattedTable();
                Console.WriteLine($"格式化表格创建结果:");
                Console.WriteLine($"  表格路径: {formattedTableResult.DocumentPath}");
                Console.WriteLine($"  是否应用样式: {formattedTableResult.IsStyled}");

                Console.WriteLine("使用辅助类的完整示例操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例操作出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 表格操作管理器辅助类
    /// </summary>
    public class TableOperationsManager
    {
        /// <summary>
        /// 简单表格结果类
        /// </summary>
        public class SimpleTableResult
        {
            /// <summary>
            /// 表格行数
            /// </summary>
            public int RowCount { get; set; }

            /// <summary>
            /// 表格列数
            /// </summary>
            public int ColumnCount { get; set; }
        }

        /// <summary>
        /// 数据表格结果类
        /// </summary>
        public class DataTableResult
        {
            /// <summary>
            /// 表格标题
            /// </summary>
            public string Title { get; set; }

            /// <summary>
            /// 数据行数
            /// </summary>
            public int DataRowCount { get; set; }
        }

        /// <summary>
        /// 格式化表格结果类
        /// </summary>
        public class FormattedTableResult
        {
            /// <summary>
            /// 文档路径
            /// </summary>
            public string DocumentPath { get; set; }

            /// <summary>
            /// 是否应用样式
            /// </summary>
            public bool IsStyled { get; set; }
        }

        /// <summary>
        /// 创建简单表格
        /// </summary>
        /// <returns>简单表格结果</returns>
        public SimpleTableResult CreateSimpleTable()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 创建表格
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            var table = document.Tables.Add(range, 3, 4);

            return new SimpleTableResult
            {
                RowCount = table.Rows.Count,
                ColumnCount = table.Columns.Count
            };
        }

        /// <summary>
        /// 创建数据表格
        /// </summary>
        /// <returns>数据表格结果</returns>
        public DataTableResult CreateDataTable()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 创建表格
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            var table = document.Tables.Add(range, 5, 3);

            // 设置标题
            table.Title = "员工信息表";

            // 填充表头
            table.Cell(1, 1).Range.Text = "姓名";
            table.Cell(1, 2).Range.Text = "部门";
            table.Cell(1, 3).Range.Text = "职位";

            // 填充数据
            string[,] data = {
                {"张三", "技术部", "工程师"},
                {"李四", "市场部", "经理"},
                {"王五", "人事部", "专员"},
                {"赵六", "财务部", "会计"}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    table.Cell(i + 2, j + 1).Range.Text = data[i, j];
                }
            }

            return new DataTableResult
            {
                Title = table.Title,
                DataRowCount = data.GetLength(0)
            };
        }

        /// <summary>
        /// 创建格式化表格
        /// </summary>
        /// <returns>格式化表格结果</returns>
        public FormattedTableResult CreateFormattedTable()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 创建表格
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            var table = document.Tables.Add(range, 4, 3);

            // 填充数据
            table.Cell(1, 1).Range.Text = "产品";
            table.Cell(1, 2).Range.Text = "销量";
            table.Cell(1, 3).Range.Text = "价格";

            table.Cell(2, 1).Range.Text = "产品A";
            table.Cell(2, 2).Range.Text = "100";
            table.Cell(2, 3).Range.Text = "50";

            table.Cell(3, 1).Range.Text = "产品B";
            table.Cell(3, 2).Range.Text = "200";
            table.Cell(3, 3).Range.Text = "30";

            // 应用样式
            table.Style = "彩色列表";

            // 保存文档
            string filePath = Path.Combine(Path.GetTempPath(), $"FormattedTable_{Guid.NewGuid()}.docx");
            document.SaveAs(filePath);

            return new FormattedTableResult
            {
                DocumentPath = filePath,
                IsStyled = !string.IsNullOrEmpty(table.Style?.ToString())
            };
        }
    }
}
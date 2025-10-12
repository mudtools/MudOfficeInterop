using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeSample
{
    /// <summary>
    /// 邮件合并模板构建器类
    /// </summary>
    public class MailMergeTemplateBuilder
    {
        private readonly IWordDocument _document;
        private readonly MailMergeHelper _mailMergeHelper;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public MailMergeTemplateBuilder(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _mailMergeHelper = new MailMergeHelper(document);
        }

        /// <summary>
        /// 创建信函模板
        /// </summary>
        /// <param name="companyName">公司名称</param>
        /// <param name="companyAddress">公司地址</param>
        /// <param name="companyPhone">公司电话</param>
        /// <returns>是否创建成功</returns>
        public bool CreateLetterTemplate(string companyName, string companyAddress, string companyPhone)
        {
            try
            {
                // 设置文档为邮件合并主文档
                _mailMergeHelper.SetMainDocumentType(WdMailMergeMainDocType.wdFormLetters);

                // 添加模板内容
                var range = _document.Range();

                // 页眉
                range.Text = $"{companyName}\n地址：{companyAddress}\n电话：{companyPhone}\n\n";
                range.Font.Bold = 1;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 日期
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddDateField(range);
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n";
                range.Font.Bold = 0;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                // 收件人地址
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "客户姓名");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "地址");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n";

                // 称呼
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "尊敬的 ";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "客户姓名");
                range.Text = " 先生/女士：\n\n";

                // 正文
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "感谢您一直以来对我们公司的支持与信任。我们很高兴地通知您，您的账户信息已更新。\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "以下是您的账户信息：\n\n";

                // 账户信息
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "客户编号：\t";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "客户编号");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "账户余额：\t";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "账户余额");
                range.Text = " 元\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "信用等级：\t";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "信用等级");
                range.Text = "\n\n";

                // 条款和条件
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "如有任何疑问，请随时与我们联系。\n\n";

                // 结尾
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "此致\n敬礼！\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = $"{companyName}客户服务部\n";

                // 更新字段
                _mailMergeHelper.UpdateAllFields();

                Console.WriteLine("信函模板创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建信函模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建标签模板
        /// </summary>
        /// <param name="columns">列数</param>
        /// <param name="rows">行数</param>
        /// <returns>是否创建成功</returns>
        public bool CreateLabelTemplate(int columns = 3, int rows = 10)
        {
            try
            {
                // 设置文档为邮件合并主文档
                _mailMergeHelper.SetMainDocumentType(WdMailMergeMainDocType.wdMailingLabels);

                // 添加标签内容
                var range = _document.Range();

                // 添加合并字段
                _mailMergeHelper.AddMergeField(range, "姓名");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "地址");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "城市");
                range.Text = " ";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "邮编");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";

                // 更新字段
                _mailMergeHelper.UpdateAllFields();

                Console.WriteLine("标签模板创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建标签模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建目录模板
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateCatalogTemplate()
        {
            try
            {
                // 设置文档为邮件合并主文档
                _mailMergeHelper.SetMainDocumentType(WdMailMergeMainDocType.wdCatalog);

                // 添加目录内容
                var range = _document.Range();

                // 添加产品信息
                range.Text = "产品目录\n\n";
                range.Font.Bold = 1;
                range.Font.Size = 16;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Font.Bold = 0;
                range.Font.Size = 12;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                // 产品详情
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "产品名称");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "描述：";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "产品描述");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "价格：￥";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "价格");
                range.Text = "\n\n";

                // 更新字段
                _mailMergeHelper.UpdateAllFields();

                Console.WriteLine("目录模板创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建目录模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建信封模板
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateEnvelopeTemplate()
        {
            try
            {
                // 设置文档为邮件合并主文档
                _mailMergeHelper.SetMainDocumentType(WdMailMergeMainDocType.wdEnvelopes);

                // 添加信封内容
                var range = _document.Range();

                // 收件人地址
                _mailMergeHelper.AddMergeField(range, "收件人姓名");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "收件人地址");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "收件人城市");
                range.Text = " ";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "收件人邮编");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n";

                // 发件人地址
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "发件人：\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "发件人姓名");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "发件人地址");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "发件人城市");
                range.Text = " ";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                _mailMergeHelper.AddMergeField(range, "发件人邮编");

                // 更新字段
                _mailMergeHelper.UpdateAllFields();

                Console.WriteLine("信封模板创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建信封模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建自定义模板
        /// </summary>
        /// <param name="templateContent">模板内容定义</param>
        /// <returns>是否创建成功</returns>
        public bool CreateCustomTemplate(List<TemplateField> templateContent)
        {
            try
            {
                var range = _document.Range();

                foreach (var field in templateContent)
                {
                    switch (field.FieldType)
                    {
                        case TemplateFieldType.Text:
                            range.Text = field.Content;
                            range.Collapse(WdCollapseDirection.wdCollapseEnd);
                            break;

                        case TemplateFieldType.MergeField:
                            _mailMergeHelper.AddMergeField(range, field.Content);
                            range.Collapse(WdCollapseDirection.wdCollapseEnd);
                            break;

                        case TemplateFieldType.DateField:
                            _mailMergeHelper.AddDateField(range, field.Content);
                            range.Collapse(WdCollapseDirection.wdCollapseEnd);
                            break;

                        case TemplateFieldType.ConditionalField:
                            _mailMergeHelper.AddConditionalField(range, field.Content);
                            range.Collapse(WdCollapseDirection.wdCollapseEnd);
                            break;

                        case TemplateFieldType.NewLine:
                            range.Text = "\n";
                            range.Collapse(WdCollapseDirection.wdCollapseEnd);
                            break;
                    }
                }

                // 更新字段
                _mailMergeHelper.UpdateAllFields();

                Console.WriteLine("自定义模板创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建自定义模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加条件文本（根据性别）
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <returns>是否添加成功</returns>
        public bool AddGenderBasedSalutation(IWordRange range)
        {
            try
            {
                // 添加条件字段，根据性别显示不同称谓
                range.Text = "{ IF { MERGEFIELD 性别 } = \"男\" \"先生\" \"女士\" }";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                Console.WriteLine("已添加基于性别的称谓条件字段");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加性别称谓条件字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加计算字段（计算折扣）
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="originalField">原字段名</param>
        /// <param name="discountRate">折扣率</param>
        /// <returns>是否添加成功</returns>
        public bool AddDiscountCalculationField(IWordRange range, string originalField, double discountRate)
        {
            try
            {
                // 添加计算字段，计算折扣价格
                range.Text = $"{{ = {{ MERGEFIELD {originalField} }} * {1 - discountRate} }}";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                Console.WriteLine($"已添加折扣计算字段，折扣率: {discountRate}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加折扣计算字段时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 模板字段类
    /// </summary>
    public class TemplateField
    {
        /// <summary>
        /// 字段类型
        /// </summary>
        public TemplateFieldType FieldType { get; set; }

        /// <summary>
        /// 字段内容
        /// </summary>
        public string Content { get; set; }
    }

    /// <summary>
    /// 模板字段类型枚举
    /// </summary>
    public enum TemplateFieldType
    {
        /// <summary>
        /// 普通文本
        /// </summary>
        Text,

        /// <summary>
        /// 合并字段
        /// </summary>
        MergeField,

        /// <summary>
        /// 日期字段
        /// </summary>
        DateField,

        /// <summary>
        /// 条件字段
        /// </summary>
        ConditionalField,

        /// <summary>
        /// 换行符
        /// </summary>
        NewLine
    }
}
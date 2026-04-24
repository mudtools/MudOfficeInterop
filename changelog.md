# OfficeInterop 项目变更日志

## 2.0.9

- 添加 IPowerPointPane 和 IPowerPointPanes 窗格接口定义
- 重构 IPowerPointDocumentWindow 接口并完善属性方法定义，删除旧实现类
- 重构 IPowerPointDocumentWindows 接口继承 IOfficeObject，添加 PpArrangeStyle 窗口排列枚举
- 添加 IWordRange.Editors 和 IWordFind.ParentRange 属性
- 添加 IWordParagraph 接口属性定义
- 统一表格边框启用属性为数值类型，重命名 PageSize 为 PaperSize、Style 为 StyleName
- 替换 ExcelOperationException 为更通用的 OfficeOperationException，统一异常处理
- 使用 TextFrame2 替代 TextFrame 设置 SmartArt 文本
- 使用 System.Drawing.Color 替代 WdColor 枚举处理图形颜色
- 示例项目统一使用 SaveAs 方法并修正参数命名规范
- 修复 WordFactory 静态类 BlankDocument 方法的实现问题

## 2.0.8

- 添加打印相关枚举和接口，重构视图接口实现
- 重构 PowerPoint 演示文稿接口继承关系并删除实现类
- 重构 PowerPoint 接口继承并简化选择接口实现
- 移除废弃的打印选项接口和实现类
- 添加 PowerPoint 文本样式和母版相关接口定义

## 2.0.7

- 更新项目版本号至 2.0.7
- 统一接口继承 IOfficeObject 以增强类型安全
- 添加大小写转换枚举并重构文本范围接口
- 添加 PowerPoint 标尺和制表位相关接口及枚举
- 重构标签接口并移除实现类
- 重构声音效果接口并移除实现类
- 移除幻灯片统计相关接口和实现类
- 添加幻灯片导航接口并重构放映窗口接口
- 添加 PowerPoint 枚举类型和接口定义
- 添加切换速度枚举并重构切换效果接口
- 添加 PpPasteDataType 枚举和 IPowerPointPlaceholders 接口
- 添加幻灯片放映相关枚举和接口
- 移动接口文件到 Core 目录并扩展功能
- 添加 PowerPoint 注释接口并重构幻灯片相关接口
- 重构动画序列和时间线接口及实现
- 重构动画序列接口并删除旧实现
- 添加 PowerPoint 动画相关枚举和接口
- 重构 IPowerPointHyperlink 接口并添加 MsoHyperlinkType 枚举
- 新增 IPowerPointGuide 相关接口和枚举
- 添加动画相关枚举并重构动画设置接口
- 重构阴影格式接口并移除实现类
- 移除冗余属性注解并重构播放设置接口
- 添加 PowerPoint 段落格式和项目符号相关枚举及接口
- 完善 PowerPoint OLE 格式接口及相关功能
- 重构 IPowerPointLineFormat 接口并删除实现类
- 重构页眉页脚接口并移动文件位置
- 更新页眉页脚接口并添加日期时间格式枚举
- 重构动作设置接口并删除实现类
- 添加 PowerPoint 相关枚举和接口定义
- 使用完整命名空间引用 PpColorSchemeIndex
- 添加新的枚举类型 PpColorSchemeIndex 和接口 IPowerPointColorFormat
- 更新接口泛型参数以包含具体 COM 类型

## 2.0.6

- 为 PowerPoint COM 属性添加包装注解并调整可为空类型
- 统一 PowerPoint 接口属性并添加 ComPropertyWrap 特性
- 优化 PowerPoint COM 属性包装和类型定义
- 为 COM 接口属性添加 ComPropertyWrap 特性
- 更新 Excel 控件接口继承以匹配 MsExcel 类型
- 更新 PowerPoint 接口继承以支持具体 COM 类型
- 统一 PowerPoint 接口继承 IOfficeObject 泛型参数
- 为接口添加泛型 MsExcel 类型参数以增强类型安全
- 为接口添加 MsExcel COM 类型参数泛型支持
- 移除打印相关枚举和 word 对话框枚举文件
- 移除过时的 Excel 枚举定义代码文件
- 优化 Excel 接口结构并增强功能支持

## 2.0.5

- 更新 Excel 操作对象库版本并优化方法参数
- 更新接口继承类型定义并移除冗余核心视图接口
- 修正 Top 属性的数据类型转换
- 更新代码生成器版本并完成枚举类型文件
- 重新组织 IWordRange 接口代码
- 重构 Word 替换接口实现
- 重构 Word 列表模板接口并移除实现类
- 更新 IWordListParagraphs 接口添加 COM 封装支持
- 重构 IWordListGalleries 接口实现
- 更新 IWordListGallery 接口并移除实现类
- 移除 Word 段落接口定义
- 实现 IWordParagraphs 接口并移除具体实现类
- 更新 IWordParagraph 接口实现并移除具体实现类
- 重构 IWordLineNumbering 接口继承和属性类型定义
- 重构 Word 框架集合接口并移除实现类
- 更新 IWordFootnotes 接口以增强脚注功能
- 调整接口属性类型和命名规范
- 移除公共图形对象和工作表相关接口
- 更新 IWordDocuments 接口实现 COM 封装和添加文档操作功能
- 重构表格样式接口实现
- 优化 Word XML 接口和形状组合接口的类型定义
- 扩展 IWordShapeRange 接口功能并移除冗余实现
- 更新 IWordShape 接口并移除冗余的 WordShape 实现类
- 重构 IWordShading 接口以使用通用 Office 对象封装
- 更新 IWordInterior 接口实现 COM 对象封装
- 添加 Value2 属性支持并更新数据透视表接口
- 更新.gitignore 文件添加 Excel 示例项目 obj 目录
- 更新 IWordInlineShapes 接口增强 Word 内嵌形状功能
- 重构标题样式和内联形状接口以支持 Office 对象封装
- 更新 IWordHeadingStyle 接口以支持 COM 对象封装
- 更新 IWordGroupShapes 接口实现
- 重构 IWordBorders 接口并移除实现类
- 重构 IWordBorder 接口并移除实现类
- 更新 Word 格式接口继承 IOfficeObject
- 重构图片格式接口实现
- 实现 IWordOLEFormat 接口并移除旧的封装类
- 更新 Word 列表相关接口实现
- 为 Word 格式接口添加 COM 属性包装器
- 重构 IWordLinkFormat 接口实现
- 更新 IWordLineFormat 接口定义并移除具体实现类
- 更新 IWordGlowFormat 接口实现 COM 对象封装
- 为接口添加 OfficeObject 继承支持
- 更新 TwoInitialCapsException 接口实现
- 优化自动更正异常接口设计并移除实现类
- 重构 IWordOtherCorrectionsException 接口实现
- 重构连字和字母自动更正异常接口
- 重构 HangulAndAlphabetException 接口实现
- 更新 IWordFirstLetterExceptions 接口并移除实现类
- 重构 IWordFirstLetterException 接口实现
- 添加 Word 自动更正条目接口属性和方法
- 重构 IWordAutoCorrectEntries 接口定义并移除实现类
- 重构 IWordAutoCorrect 接口实现

## 2.0.4

- 升级版本号至 2.0.4 并更新依赖包
- 更新接口继承以支持泛型类型参数
- 更新 Word 接口继承类型以包含具体 COM 对象类型
- 将 ComCollectionWrap 属性更改为 ComObjectWrap 并添加泛型 COM 类型支持
- 更新 IWordVariables 接口定义并移除实现类
- 重构 Word 变量接口实现
- 重构 Word 文档接口实现
- 更新 IWordSubdocuments 接口添加 COM 封装和可空类型支持
- 重构子文档接口实现并添加 COM 封装特性
- 更新 IWordSources 接口实现并移除具体实现类
- 重构智能标记类型接口和源接口实现
- 重构智能标记识别程序接口
- 更新 IWordSmartTagActions 接口添加 COM 封装和额外功能
- 更新接口继承关系并移除实现类
- 更新 Word 评论相关接口实现
- 添加 WordProofreadingErrors 接口的 COM 封装支持
- 重构 Word 可读性统计和最近文档接口
- 为 Word 接口添加 COM 封装特性和 Application 属性说明
- 添加数据透视表缓存属性和 VBE 组件功能
- 优化代码格式和类型定义

## 2.0.3

- 更新版本号至 2.0.3
- 重构 Word 对话框和对话框集合的具体实现类
- 重构 Word 页面接口实现
- 更新 Word 接口定义并移除实现类
- 更新 IWordPageNumbers 接口并移除实现类
- 更新 ServiceCodeGenerator 版本并优化枚举类型接口
- 移除内容控件和表单域相关接口及实现类
- 重构文件转换器接口实现
- 完成 VBE 对象接口实现并添加类型注释
- 更新 IWordEmailOptions 接口定义并移除实现类
- 为 AutoCorrect 接口添加 COM 属性包装器
- 移除辅助接口定义
- 重构 IWordEditor 和 IWordEditors 接口实现
- 移除 Word 下拉列表功能相关接口和实现
- 更新字典接口实现并添加加载器重载
- 为 Word 接口添加 Office 对象支持并调整属性类型
- 更新 IWordCustomProperties 接口添加 COM 封装支持和加载器
- 添加 Word 自定义属性接口功能
- 添加项目英文版本 README 文件
- 更新 IWordContentControls 接口定义并移除实现类
- 为 IWordContentControlListEntries 接口添加 COM 封装支持并移除实现类
- 更新 Word 操作对象接口以支持更好的 COM 对象管理
- 更新 IWordContentControl 接口继承和属性可空性
- 更新 IWordConflicts 接口定义并移除实现类
- 重构 IWordConflict 接口以支持 COM 对象封装
- 更新 Word 权限接口实现并移除旧的实现类
- 移除 IWordBrowser 接口
- 移除 WordCheckBox 相关实现
- 更新接口继承结构并移除实现类
- 更新 IWordBookmarks 接口添加 COM 封装支持
- 更新目录和书签接口实现
- 重构 AutoCaption 接口实现
- 调整 IWordAdjustments 接口实现并移除具体实现类
- 更新 IWordAddIn 和 IWordAddIns 接口实现
- 为 IWordAddIn 接口添加 ComObjectWrap 特性
- 更新 IWordPage 接口添加页面功能和属性
- 修复 Directory.Build.props 文件中的 Mud.ServiceCodeGenerator 包版本
- 更新 IExcelRanges 接口定义并移除实现类
- 更新 Excel AddIns 接口实现并改进类型安全
- 为 IExcelAddIn 接口添加 Application 和 Parent 属性
- 添加用户访问权限接口和改进工作表接口
- 为 Excel 接口添加 Office 对象继承
- 更新 IExcelWindow 接口定义并移除实现类
- 重构受保护视图窗口接口并移除实现类
- IExcelRange 接口实现 IOfficeObject 接口并添加 LoadFromObject 方法
- 添加 ExcelQueryTable 接口实现和属性重构
- 替换 ReturnValueConvert 为 ValueConvert 属性并优化接口
- 更新 IExcelWorkbook 接口类型定义为包装接口
- 为 IExcelCellFormat 接口添加 IOfficeObject 继承
- 更新 IExcelWorkbooks 接口并移除实现类
- 移除 PrintPreview 相关代码并调整 PageSetup 属性
- 重构 Excel 页面设置接口并添加 COM 封装属性
- 重构 Excel 评论相关接口和实现
- 更新 IExcelCellFormat 接口定义并移除实现类
- 移除高级图表功能说明文档
- 为 Excel 控件接口添加 Office 对象继承
- 为 Excel 格式化样式接口添加 OfficeObject 继承
- 重构 IExcelTickLabels 接口并移除实现类
- 为 Excel 格式接口添加 COM 属性包装
- 修复 CoreRange 格式设置问题并优化 excelStyles 接口设计
- 更新 IExcelStyle 接口定义并移除实现类
- 移除 SmartTagActions 和 SmartTags 实现类并更新接口定义
- 添加 Excel 智能标记操作接口实现
- 更新图片和形状接口定义添加 COM 封装属性
- 更新 Excel 接口的可空类型定义并优化分页符接口
- 更新 IExcelGroupShapes 接口定义
- 移除 IExcelFreeformBuilder 接口并调整 ExcelGridlines 实现
- 更新 IExcelErrorCheckingOptions 接口增强错误检查功能
- 为 ExcelDataBarBorder 接口添加 COM 对象封装支持
- 更新颜色刻度条件接口以支持空值和字符串加载
- 将 IWordApplication 属性改为可空类型
- 为 Excel 格式接口添加 IOfficeObject 继承并更新可空类型
- 更新接口定义以支持可空类型
- 为 Excel 接口添加 IOfficeObject 继承
- 优化对象扩展和查找替换功能
- 添加 Office COM 对象封装器创建功能
- 添加 IOfficeObject 接口
- 为 Excel 添加 IOfficeObject 接口并支持空值返回
- 为 XML 接口添加 IOfficeObject 继承并实现工厂创建方法
- 为 SmartArt 相关接口添加 IOfficeObject 类型继承
- 为 Office 操作对象接口添加通用对象类型
- 为 Office 格式接口添加 IOfficeObject 继承
- 更新 Word 接口定义和实现
- 添加 Word 数学公式排序示例项目的构建输出目录
- 移除 IWordDataLabels 接口中的数据标签相关属性
- 更新 Excel 示例项目的.gitignore 文件
- 更新 IWordFields 接口并移除实现类
- 更新 Word 查找接口和表格样式属性
- 为 Word 文本接口添加 COM 封装特性和辅助生成器特性

## 2.0.1

- 重构 Excel COM 接口封装类
- 重构 Excel 条件格式接口定义
- 更新接口定义以增强空值处理与类型安全性
- 更新 Excel 表格格式接口并增强功能
- 完成 Excel 趋势线集合接口定义以增强空值处理能力
- 完成趋势线接口定义并强化 COM 对象封装
- 为 IExcelUpBars 接口添加 COM 属性与新属性
- 为 ISlicerPivotTables 接口添加 COM 封装特性并优化空值处理
- 更新项目版本号并调整打包配置属性
- 更新数据透视表相关接口返回值可空性
- 统一使用 InternalComObject 访问底层 COM 对象
- 重构图表相关接口和实现类
- 扩展图表组合接口功能
- 完成 Excel 坐标轴刻度接口定义
- 完成 Excel 三维格式接口定义
- 为 ExcelTextEffectFormat 接口添加 COM 属性包装器
- 完成 Excel 阴影格式接口定义并移除冗余实现
- 为 Excel 图片格式接口添加 COM 封装特性
- 重构 Excel 图表相关接口和实现
- 完成 Excel 图表图例接口定义
- 移除部分接口中废弃的实现
- 完善错误处理机制设计
- 增加全局的 DisposableList 资源释放对象
- 重构 Excel 工作簿数据连接接口实现
- 优化 Excel 数据透视表接口和文档连接接口
- 完成 Excel ODBC 和 OLEDB 连接接口功能
- 更新 Excel 验证接口并优化参数转换方法
- 重构 Excel 表格对象和文档连接接口
- 重构 Excel 模型连接接口实现
- 重构 Excel 工作表行接口并增强功能
- 重构 Excel 排序接口并更新项目名称
- 重构 Excel 切片器接口并增强功能
- 更新 Mud.ServiceCodeGenerator 包版本
- 移除 IExcelSlicer 接口中的 Columns 属性和 IExcelSlicerCache 接口中的 IsPivot 属性
- 为 IExcelSlicer 添加 Copy 和 Cut 方法并完善接口定义
- 为多个 Excel COM 接口添加 Application 和 Parent 属性
- 更新 Excel 最近文档接口定义并调整实现
- 优化 Excel 页眉页脚接口实现
- 为 IExcelGraphic 接口添加 ColorType 属性支持
- 更新 Excel 接口定义与实现，增强 COM 封装一致性与可维护性

## 1.2.2

- 添加自动生成代码所需的特性
- 删除不需要的文件
- 完成 Excel 样式集合的接口的资源释放功能
- 完成 Excel 页面集合的接口的资源释放
- 完成 Excel 筛选器接口的资源释放功能
- 完成 ExcelErrors 资源释放
- 删除无效的文件
- 重命名 AssemblyInfo 文件
- 完成测试代码
- 删除演示代码中文本对话框设置功能
- 升级示例文件的引用程序集
- 添加 IOfficePermission 封装接口
- 完成 IWordFillFormat 接口封装
- 更新包的引用和依赖
- 解决 Excel 示例程序中的编译错误
- IWordDocument 封装接口添加 SaveAs 方法
- IWordShapes 二次封装对象添加 AddTextEffect 方法
- IExcelShapes 二次封装对象添加 AddChart2 方法
- 完成 Excel 图表组件的相关二次封装功能
- 修改 Excel 示例代码
- 添加 Excel 示例代码
- 添加 Excel 相关的功能文件
- 修正文件夹合并示例错误
- 完成文件夹合并示例
- 移除多余的开发信息
- 添加 IExcelChartGroup 二次封装对象
- 将 IExcelChart 接口的 Axes 属性调用改为 Axes 方法
- 添加 word 相关文档
- 发布 1.1.9 版本
- 添加 IExcelRangeCharacters 二次封装接口
- 完善 IWordFont 缺失的属性及方法
- 优化 IExcelHyperlinks 接口的资源管理
- 完成 IExcelHPageBreak 接口
- 优化 IExcelHPageBreaks 接口
- 优化 ExcelHPageBreak 空对象处理机制
- 完善 IExcelFormatCondition 接口缺失的属性及方法
- 完成 IExcelColorScaleCriteria 二次封装接口
- 优化 IExcelBorders 二次封装接口
- 完成 IExcelValidation 接口
- 优化 IWordBorders 二次封装接口的参数转换方式
- 优化 Excel 集合类 COM 组件的资源释放功能
- 添加 IExcelAxisTitle 二次封装接口
- 发布 1.1.8
- 完成 IExcelDataBar 二次封装接口
- 完成 Excel 文档内容
- 添加 API 文档生成组件
- 完成系统异常处理架构
- 添加全局的 DisposableList 资源释放对象
- 修改 gitignore
- 完成 IExcelDrawing 二次封装接口
- 完成 IExcelDataTable 封装代码
- 完成 IExcelChartTitle 缺失的属性
- 完成 IExcelChart 二次封装接口
- 完成 IExcelSmartTag 二次封装接口
- 完成 IExcelStyle 二次封装接口
- 完成 IExcelBorder 二次封装接口代码
- 添加 IOfficeSmartArt 二次封装接口
- 完成 IExcelShapeRange 缺失的属性及方法
- 添加 IExcelGroupObject 二次封装接口

## 1.1.6

- IExcelRange 添加 CopyPicture 方法
- 将 IExcelComSheets 接口的 EnumerateSheets 方法重命名为 Items
- IExcelComSheets 添加 this 索引器

## 1.1.5

- 完成 IExcelShape 二次封装接口缺失的属性
- 完成 IExcelPicture 接口
- 完成 IExcelTextFrame 缺失的属性
- 完成 IExcelSortFields 二次封装接口的实现
- 修改系统枚举转换方法
- 修改打印预览组件
- 加入 Excel 中剪切板对象封装相关类
- 添加查询对象 IExcelQueryTable 及相关的封装接口
- 添加 IExcelBorders 缺失的属性
- IExcelBorder 接口添加 ColorIndex 属性
- 重构 IExcelComSheets 接口

## 1.1.4

- 整理 ICommonWorksheet 的公共属性及方法
- IExcelChart 二次封装对象绑定事件处理
- 为 ICommonWorksheet 接口添加公共的 Protect 方法
- 整理 IExcelRange 的 SpecialCells 方法签名
- 整理 IExcelWorksheet 的 AutoFilter 属性
- IExcelWindow 二次封装对象添加 ActiveCell、ActiveChart、ActivePane、ActiveSheet 和 ActiveWorksheet 属性
- 修复编译错误
- 完成 IExcelPane 二次封装对象
- IExcelShapeRange 接口添加 ScaleWidth、ScaleHeight 方法
- ICoreRange 添加 AutoFormat 方法和 AutoOutline 方法
- 整理 IExcelWorkbook 的 PivotCaches 方法返回对象
- ICoreRange 接口添加 Replace 方法
- 修改 ICoreRange 的 BorderAround 方法签名
- 修改 ICoreRange 的 Find 方法签名
- 完成接口注释
- 完善缺失的属性注释内容
- 优化 ExcelCellFormat 属性
- 处理 IWordPageSetup 接口代码
- 完成 IWordPageSetup 组件二次封装
- 继续完成 IWordPageSetup 封装属性
- 完成 IWordPageSetup 二次封装接口
- IExcelRange 添加 Parse 方法
- 整理 IExcelRange 的 GetAddressLocal、GetAddress 方法签名
- 完成 ExcelCommonSheets 的异常处理
- 整理 IExcelComments 的 Add 方法
- 完成 IExcelWindow 二次封装对象

## 1.1.1

- 完成 IWordDocument 二次封装对象
- 完成 IWordWindow 二次封装接口
- 添加 OLEFormat 类封装

## 1.1.0

- 处理 WordRange 二次封装属性及资源释放
- 添加 IWordHeadersFooters 封装类
- 完成 IWordApplication 接口的封装
- 完善部分类的注释内容
- 处理 WordApplication 接口及实现的代码
- 完成枚举注释
- 完成 IOfficePickerDialog 二次封装对象
- 继续完成 WordApplication 相关的附属属性封装
- 完成 Office Search 组件封装
- 添加文件搜索公共接口封装
- 整理 Excel 项目文件结构
- 处理 Excel 项目中的枚举文件
- 整理 MudTools.OfficeInterop.Word 项目中文件从属
- 整理 Word 项目中的枚举
- 完善 WordApplication 封装类
- 添加自定义组合键 KeyBindings 二次封装
- 完成 Application 二次封装组件
- 解决现有封装类中存在的编译错误
- 添加 IWordEditor、IWordConflict 封装接口
- 完成 IWordRange 二次封装的属性及方法
- 完善完成 WordTable、WordBorders 缺失的属性及方法
- 整理 Shape 的从属文件夹
- 添加 Word 对象错误集合的封装实现类
- 添加 Word 文档中的一个子文档（Subdocument）的封装接口
- 添加 Word 文档中修订（Revision）的封装接口
- 添加 Word FormField 封装代码
- 添加 Word 单元格外相关二次封装代码
- 完成 word 文本处理相关的封装接口
- 整理图形相关的封装接口
- 整理文本相关类的文件结构
- 完成 Word Chart Com 组件的二次封装
- 添加代码版权信息
- 修正添加 Application 属性引起的错误
- 添加 WordShape 的二次封装代码
- 添加用于操作 Word 表格样式 TableStyle 组件
- 添加文档段落格式相关的缩进字符单元内内容解析文件
- 添加用于操作 Word 文档中段落的格式设置 ParagraphFormat COM 组件相关类
- 添加 AutoCorrectEntries 自动更正条目 COM 组件封装代码
- 添加 IExcelRange 高级应用文档
- 添加 Word 模板相关二次封装代码
- 完成接口文档
- 添加类的描述信息
- 完成打印设置功能说明文档
- 更新操作指南文档

## 1.0.7

- 修复 IExcelPicture 封装的错误
- 为 ICommonWorksheet 添加 ParentName 属性
- 解决 MudTools.OfficeInterop.Word 项目包与 PPT 项目包冲突的问题
- 修改系统发布目标平台
- 加入 IExcelErrors 接口
- 修改项目的 API 文档内容
- 修改编译配置
- 加入 AssemblyName 描述信息
- 加入项目打包信息
- 分别为 MudTools.OfficeInterop、MudTools.OfficeInterop.Excel、MudTools.OfficeInterop.PowerPoint、MudTools.OfficeInterop.Word 项目编写 README.md 文件
- 删除项目中大量不需要的类
- 添加 README.md 文件

## 1.0.0

- 首次提交代码
- 删除源码管理不需要的文件

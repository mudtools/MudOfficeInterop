//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel WorksheetFunction 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.WorksheetFunction 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelWorksheetFunction : IOfficeObject<IExcelWorksheetFunction, MsExcel.WorksheetFunction>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 WorksheetFunction 对象的父对象（通常是 Application）
    /// 对应 WorksheetFunction.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取 WorksheetFunction 对象所在的Application对象
    /// 对应 WorksheetFunction.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 数学和三角函数 

    /// <summary>
    /// 返回数字的反余弦值
    /// 对应 WorksheetFunction.Acos 方法
    /// </summary>
    /// <param name="number">介于 -1 到 1 之间的数字</param>
    /// <returns>反余弦值</returns>
    double? Acos(double number);

    /// <summary>
    /// 返回数字的反正弦值
    /// 对应 WorksheetFunction.Asin 方法
    /// </summary>
    /// <param name="number">介于 -1 到 1 之间的数字</param>
    /// <returns>反正弦值</returns>
    double? Asin(double number);

    /// <summary>
    /// 返回直角三角形的反正切值
    /// 对应 WorksheetFunction.Atan2 方法
    /// </summary>
    /// <param name="xNum">直角三角形的底边长度</param>
    /// <param name="yNum">直角三角形的高</param>
    /// <returns>反正切值</returns>
    double? Atan2(double xNum, double yNum);

    /// <summary>
    /// 返回数字的向上舍入值（远离零的方向）
    /// 对应 WorksheetFunction.Ceiling 方法
    /// </summary>
    /// <param name="number">要舍入的数字</param>
    /// <param name="significance">用作舍入基准的倍数</param>
    /// <returns>向上舍入值</returns>
    double? Ceiling(double number, double significance);

    /// <summary>
    /// 返回数字组合数
    /// 对应 WorksheetFunction.Combin 方法
    /// </summary>
    /// <param name="number">项目的总数</param>
    /// <param name="numberChosen">每一组合中项目的数量</param>
    /// <returns>组合数</returns>
    double? Combin(double number, double numberChosen);

    /// <summary>
    /// 返回数字的双曲余弦值
    /// 对应 WorksheetFunction.Cosh 方法
    /// </summary>
    /// <param name="number">任意实数</param>
    /// <returns>双曲余弦值</returns>
    double? Cosh(double number);

    /// <summary>
    /// 将弧度转换为度
    /// 对应 WorksheetFunction.Degrees 方法
    /// </summary>
    /// <param name="angle">以弧度为单位的角度</param>
    /// <returns>度数</returns>
    double? Degrees(double angle);

    /// <summary>
    /// 返回小于或等于数字的最大整数
    /// 对应 WorksheetFunction.Floor 方法
    /// </summary>
    /// <param name="number">要舍入的数字</param>
    /// <param name="significance">用作舍入基准的倍数</param>
    /// <returns>向下舍入值</returns>
    double? Floor(double number, double significance);

    /// <summary>
    /// 返回数字的阶乘
    /// 对应 WorksheetFunction.Fact 方法
    /// </summary>
    /// <param name="number">要计算其阶乘的非负数</param>
    /// <returns>阶乘</returns>
    double? Fact(double number);

    /// <summary>
    /// 返回数字的双精度阶乘
    /// 对应 WorksheetFunction.FactDouble 方法
    /// </summary>
    /// <param name="number">要计算其双精度阶乘的数字</param>
    /// <returns>双精度阶乘</returns>
    double? FactDouble(double number);

    /// <summary>
    /// 返回数字的自然对数
    /// 对应 WorksheetFunction.Ln 方法
    /// </summary>
    /// <param name="number">要计算其自然对数的正实数</param>
    /// <returns>自然对数</returns>
    double? Ln(double number);

    /// <summary>
    /// 返回数字的指定底数的对数
    /// 对应 WorksheetFunction.Log 方法
    /// </summary>
    /// <param name="number">要计算其对数的正实数</param>
    /// <param name="baseNumber">对数的底数</param>
    /// <returns>指定底数的对数</returns>
    double? Log(double number, object? baseNumber = null);

    /// <summary>
    /// 返回数字以 10 为底的对数
    /// 对应 WorksheetFunction.Log10 方法
    /// </summary>
    /// <param name="number">要计算其常用对数的正实数</param>
    /// <returns>以 10 为底的对数</returns>
    double? Log10(double number);

    /// <summary>
    /// 返回数字的四舍五入值
    /// 对应 WorksheetFunction.Round 方法
    /// </summary>
    /// <param name="number">要舍入的数字</param>
    /// <param name="numDigits">小数位数</param>
    /// <returns>四舍五入值</returns>
    double? Round(double number, int numDigits);

    /// <summary>
    /// 返回数字的向上舍入值（远离零的方向）
    /// 对应 WorksheetFunction.RoundUp 方法
    /// </summary>
    /// <param name="number">要舍入的数字</param>
    /// <param name="numDigits">小数位数</param>
    /// <returns>向上舍入值</returns>
    double? RoundUp(double number, int numDigits);

    /// <summary>
    /// 返回数字的向下舍入值（朝零的方向）
    /// 对应 WorksheetFunction.RoundDown 方法
    /// </summary>
    /// <param name="number">要舍入的数字</param>
    /// <param name="numDigits">小数位数</param>
    /// <returns>向下舍入值</returns>
    double? RoundDown(double number, int numDigits);

    /// <summary>
    /// 返回数字的双曲正弦值
    /// 对应 WorksheetFunction.Sinh 方法
    /// </summary>
    /// <param name="number">任意实数</param>
    /// <returns>双曲正弦值</returns>
    double? Sinh(double number);

    /// <summary>
    /// 返回数字的双曲正切值
    /// 对应 WorksheetFunction.Tanh 方法
    /// </summary>
    /// <param name="number">任意实数</param>
    /// <returns>双曲正切值</returns>
    double? Tanh(double number);

    /// <summary>
    /// 将度转换为弧度
    /// 对应 WorksheetFunction.Radians 方法
    /// </summary>
    /// <param name="angle">以度为单位的角度</param>
    /// <returns>弧度值</returns>
    double? Radians(double angle);

    /// <summary>
    /// 返回数字的幂
    /// 对应 WorksheetFunction.Power 方法
    /// </summary>
    /// <param name="number">底数</param>
    /// <param name="power">指数</param>
    /// <returns>幂</returns>
    double? Power(double number, double power);
    #endregion

    #region 统计函数
    /// <summary>
    /// 返回其参数的平均值（算术平均值）
    /// 对应 WorksheetFunction.Average 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>平均值</returns>
    double? Average(object arg1, params object[] args);

    /// <summary>
    /// 返回数字参数的个数
    /// 对应 WorksheetFunction.Count 方法
    /// </summary>
    /// <param name="arg1">第一个值</param>
    /// <param name="args">其他值</param>
    /// <returns>数字参数的个数</returns>
    double? Count(object arg1, params object[] args);

    /// <summary>
    /// 计算包含数字的单元格以及参数列表中数字的个数
    /// 对应 WorksheetFunction.CountA 方法
    /// </summary>
    /// <param name="arg1">第一个值</param>
    /// <param name="args">其他值</param>
    /// <returns>非空值的个数</returns>
    double? CountA(object arg1, params object[] args);

    /// <summary>
    /// 返回数据集中的第 k 个最大值
    /// 对应 WorksheetFunction.Large 方法
    /// </summary>
    /// <param name="array">要从中查找第 k 个最大值的数组或数据区域</param>
    /// <param name="k">返回值在数组中的位置（从大到小排）</param>
    /// <returns>第 k 个最大值</returns>
    double? Large(object array, double k);

    /// <summary>
    /// 返回数据集中的第 k 个小值
    /// 对应 WorksheetFunction.Small 方法
    /// </summary>
    /// <param name="array">要从中查找第 k 个小值的数组或数据区域</param>
    /// <param name="k">返回值在数组中的位置（从小到大排）</param>
    /// <returns>第 k 个小值</returns>
    double? Small(object array, double k);

    /// <summary>
    /// 返回其参数中的最大值
    /// 对应 WorksheetFunction.Max 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>最大值</returns>
    double? Max(object arg1, params object[] args);

    /// <summary>
    /// 返回其参数中的最小值
    /// 对应 WorksheetFunction.Min 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>最小值</returns>
    double? Min(object arg1, params object[] args);

    /// <summary>
    /// 返回其参数的乘积
    /// 对应 WorksheetFunction.Product 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>乘积</returns>
    double? Product(object arg1, params object[] args);

    /// <summary>
    /// 返回其参数的总和
    /// 对应 WorksheetFunction.Sum 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>总和</returns>
    double? Sum(object arg1, params object[] args);


    /// <summary>
    /// 返回标准偏差
    /// 对应 WorksheetFunction.StDev 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>标准偏差</returns>
    double? StDev(object arg1, params object[] args);

    /// <summary>
    /// 返回基于样本总体的标准偏差
    /// 对应 WorksheetFunction.StDevP 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>总体标准偏差</returns>
    double? StDevP(object arg1, params object[] args);

    /// <summary>
    /// 返回方差
    /// 对应 WorksheetFunction.Var 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>方差</returns>
    double? Var(object arg1, params object[] args);

    /// <summary>
    /// 返回基于样本总体的方差
    /// 对应 WorksheetFunction.VarP 方法
    /// </summary>
    /// <param name="arg1">第一个数字</param>
    /// <param name="args">其他数字</param>
    /// <returns>总体方差</returns>
    double? VarP(object arg1, params object[] args);
    #endregion

    #region 文本函数
    /// <summary>
    /// 将数字转换为文本，并以指定的小数位数进行格式化
    /// 对应 WorksheetFunction.Fixed 方法
    /// </summary>
    /// <param name="number">要设置格式的数字</param>
    /// <param name="decimals">小数位数</param>
    /// <param name="noCommas">是否不显示逗号</param>
    /// <returns>格式化后的文本</returns>
    string Fixed(double number, object decimals = null, object noCommas = null);

    /// <summary>
    /// 删除文本中的所有空格（除了单词之间的单个空格）
    /// 对应 WorksheetFunction.Trim 方法
    /// </summary>
    /// <param name="text">要删除其中空格的文本</param>
    /// <returns>去除多余空格的文本</returns>
    string Trim(string text);

    /// <summary>
    /// 在文本中查找另一个文本字符串（区分大小写），并从文本字符串的第一个字符开始返回该字符串的起始位置编号
    /// 对应 WorksheetFunction.Find 方法
    /// </summary>
    /// <param name="findText">要查找的文本</param>
    /// <param name="withinText">要在其中查找 FindText 参数的文本</param>
    /// <param name="startNum">withinText 参数中从第几个字符开始查找</param>
    /// <returns>起始位置编号</returns>
    double? Find(string findText, string withinText, object? startNum = null);

    /// <summary>
    /// 返回替换指定字符数的文本字符串
    /// 对应 WorksheetFunction.Replace 方法
    /// </summary>
    /// <param name="oldText">包含要替换的字符的文本字符串</param>
    /// <param name="startNum">oldText 中要替换的第一个字符的位置</param>
    /// <param name="numChars">oldText 中要替换的字符数</param>
    /// <param name="newText">将用于替换 oldText 中字符的文本</param>
    /// <returns>替换后的文本</returns>
    string Replace(string oldText, double startNum, double numChars, string newText);

    /// <summary>
    /// 重复文本指定的次数
    /// 对应 WorksheetFunction.Rept 方法
    /// </summary>
    /// <param name="text">要重复的文本</param>
    /// <param name="numberTimes">文本重复的次数</param>
    /// <returns>重复后的文本</returns>
    string Rept(string text, double numberTimes);

    /// <summary>
    /// 返回文本字符串的子字符串
    /// 对应 WorksheetFunction.Substitute 方法
    /// </summary>
    /// <param name="text">要替换其中字符的文本</param>
    /// <param name="oldText">要替换的旧文本</param>
    /// <param name="newText">用于替换 OldText 的新文本</param>
    /// <param name="instanceNum">指定要用 newText 替换 oldText 的第几次出现</param>
    /// <returns>替换后的文本</returns>
    string Substitute(string text, string oldText, string newText, object? instanceNum = null);
    #endregion

    #region 日期和时间函数   

    /// <summary>
    /// 返回两个日期之间的天数
    /// 对应 WorksheetFunction.Days360 方法
    /// </summary>
    /// <param name="startDate">计算期间的开始日期</param>
    /// <param name="endDate">计算期间的结束日期</param>
    /// <param name="method">指示在计算中是采用美国方法 (NASD) 还是欧洲方法</param>
    /// <returns>天数</returns>
    double? Days360(object startDate, object endDate, object? method = null);

    /// <summary>
    /// 返回某日期所对应的星期数
    /// 对应 WorksheetFunction.Weekday 方法
    /// </summary>
    /// <param name="serialNumber">要查找其星期数的日期</param>
    /// <param name="returnType">用于确定返回值类型的数字</param>
    /// <returns>星期数</returns>
    double? Weekday(object serialNumber, object? returnType = null);

    /// <summary>
    /// 返回某个日期是当年的第几周
    /// 对应 WorksheetFunction.WeekNum 方法
    /// </summary>
    /// <param name="serialNumber">要确定其位于第几周的日期</param>
    /// <param name="returnType">一个数字，确定每周从哪一天开始</param>
    /// <returns>周数</returns>
    double? WeekNum(object serialNumber, object? returnType = null);
    #endregion

    #region 查找和引用函数
    /// <summary>
    /// 使用列索引从数组中选择值
    /// 对应 WorksheetFunction.Index 方法
    /// </summary>
    /// <param name="array">单元格区域或数组常量</param>
    /// <param name="rowNum">选择数组中的某行</param>
    /// <param name="columnNum">选择数组中的某列</param>
    /// <returns>选定的值</returns>
    object Index(object array, double rowNum, object? columnNum = null);

    /// <summary>
    /// 在单元格区域中查找值，然后返回该区域中指定单元格的值
    /// 对应 WorksheetFunction.HLookup 方法
    /// </summary>
    /// <param name="lookupValue">要在表格第一行中查找的值</param>
    /// <param name="tableArray">信息表</param>
    /// <param name="rowIndexNum">tableArray 中待返回值的行序号</param>
    /// <param name="rangeLookup">指定查找方式</param>
    /// <returns>查找到的值</returns>
    object HLookup(object lookupValue, object tableArray, double rowIndexNum, object? rangeLookup = null);

    /// <summary>
    /// 在单元格区域中查找值，然后返回该区域中指定单元格的值
    /// 对应 WorksheetFunction.VLookup 方法
    /// </summary>
    /// <param name="lookupValue">要在表格第一列中查找的值</param>
    /// <param name="tableArray">信息表</param>
    /// <param name="columnIndexNum">tableArray 中待返回值的列序号</param>
    /// <param name="rangeLookup">指定查找方式</param>
    /// <returns>查找到的值</returns>
    object VLookup(object lookupValue, object tableArray, double columnIndexNum, object? rangeLookup = null);

    #endregion

    #region 工程函数
    /// <summary>
    /// 将二进制数转换为十进制数
    /// 对应 WorksheetFunction.Bin2Dec 方法
    /// </summary>
    /// <param name="number">要转换的二进制数</param>
    /// <returns>十进制数</returns>
    string Bin2Dec(string number);

    /// <summary>
    /// 将二进制数转换为十六进制数
    /// 对应 WorksheetFunction.Bin2Hex 方法
    /// </summary>
    /// <param name="number">要转换的二进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>十六进制数</returns>
    string Bin2Hex(string number, object? places = null);

    /// <summary>
    /// 将二进制数转换为八进制数
    /// 对应 WorksheetFunction.Bin2Oct 方法
    /// </summary>
    /// <param name="number">要转换的二进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>八进制数</returns>
    string Bin2Oct(string number, object? places = null);

    /// <summary>
    /// 将十进制数转换为二进制数
    /// 对应 WorksheetFunction.Dec2Bin 方法
    /// </summary>
    /// <param name="number">要转换的十进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>二进制数</returns>
    string Dec2Bin(double number, object? places = null);

    /// <summary>
    /// 将十进制数转换为十六进制数
    /// 对应 WorksheetFunction.Dec2Hex 方法
    /// </summary>
    /// <param name="number">要转换的十进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>十六进制数</returns>
    string Dec2Hex(double number, object? places = null);

    /// <summary>
    /// 将十进制数转换为八进制数
    /// 对应 WorksheetFunction.Dec2Oct 方法
    /// </summary>
    /// <param name="number">要转换的十进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>八进制数</returns>
    string Dec2Oct(double number, object? places = null);

    /// <summary>
    /// 将十六进制数转换为二进制数
    /// 对应 WorksheetFunction.Hex2Bin 方法
    /// </summary>
    /// <param name="number">要转换的十六进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>二进制数</returns>
    string Hex2Bin(string number, object? places = null);

    /// <summary>
    /// 将十六进制数转换为十进制数
    /// 对应 WorksheetFunction.Hex2Dec 方法
    /// </summary>
    /// <param name="number">要转换的十六进制数</param>
    /// <returns>十进制数</returns>
    string Hex2Dec(string number);

    /// <summary>
    /// 将十六进制数转换为八进制数
    /// 对应 WorksheetFunction.Hex2Oct 方法
    /// </summary>
    /// <param name="number">要转换的十六进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>八进制数</returns>
    string Hex2Oct(string number, object? places = null);

    /// <summary>
    /// 将八进制数转换为二进制数
    /// 对应 WorksheetFunction.Oct2Bin 方法
    /// </summary>
    /// <param name="number">要转换的八进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>二进制数</returns>
    string Oct2Bin(string number, object? places = null);

    /// <summary>
    /// 将八进制数转换为十进制数
    /// 对应 WorksheetFunction.Oct2Dec 方法
    /// </summary>
    /// <param name="number">要转换的八进制数</param>
    /// <returns>十进制数</returns>
    string Oct2Dec(string number);

    /// <summary>
    /// 将八进制数转换为十六进制数
    /// 对应 WorksheetFunction.Oct2Hex 方法
    /// </summary>
    /// <param name="number">要转换的八进制数</param>
    /// <param name="places">要使用的字符数</param>
    /// <returns>十六进制数</returns>
    string Oct2Hex(string number, object? places = null);
    #endregion

    #region 信息函数
    /// <summary>
    /// 如果值为错误值 #N/A，则返回 TRUE
    /// 对应 WorksheetFunction.IsNA 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为 #N/A 错误</returns>
    bool? IsNA(object value);

    /// <summary>
    /// 如果值为任何错误值，则返回 TRUE
    /// 对应 WorksheetFunction.IsErr 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为错误值</returns>
    bool? IsErr(object value);

    /// <summary>
    /// 如果值为逻辑值，则返回 TRUE
    /// 对应 WorksheetFunction.IsLogical 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为逻辑值</returns>
    bool? IsLogical(object value);


    /// <summary>
    /// 如果值为数字，则返回 TRUE
    /// 对应 WorksheetFunction.IsNumber 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为数字</returns>
    bool? IsNumber(object value);

    /// <summary>
    /// 如果值为文本，则返回 TRUE
    /// 对应 WorksheetFunction.IsText 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为文本</returns>
    bool? IsText(object value);

    /// <summary>
    /// 如果值为偶数，则返回 TRUE
    /// 对应 WorksheetFunction.IsEven 方法
    /// </summary>
    /// <param name="number">要检验的值</param>
    /// <returns>是否为偶数</returns>
    bool? IsEven(object number);

    /// <summary>
    /// 如果值为奇数，则返回 TRUE
    /// 对应 WorksheetFunction.IsOdd 方法
    /// </summary>
    /// <param name="number">要检验的值</param>
    /// <returns>是否为奇数</returns>
    bool? IsOdd(object number);

    /// <summary>
    /// 如果值为公式产生的错误值，则返回 TRUE
    /// 对应 WorksheetFunction.IsError 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为错误值</returns>
    bool? IsError(object value);

    /// <summary>
    /// 如果值是非文本值，则返回 TRUE
    /// 对应 WorksheetFunction.IsNonText 方法
    /// </summary>
    /// <param name="value">要检验的值</param>
    /// <returns>是否为非文本值</returns>
    bool? IsNonText(object value);
    #endregion

    #region 逻辑函数
    /// <summary>
    /// 返回数组的转置
    /// 对应 WorksheetFunction.Transpose 方法
    /// </summary>
    /// <param name="arg1">要转置的数组或区域</param>
    /// <returns>转置后的数组</returns>
    object Transpose(object? arg1);

    /// <summary>
    /// 如果逻辑值参数均为 TRUE，则返回 TRUE
    /// 对应 WorksheetFunction.And 方法
    /// </summary>
    /// <param name="arg1">第一个逻辑值</param>
    /// <param name="args">其他逻辑值</param>
    /// <returns>逻辑与结果</returns>
    bool? And(object arg1, params object[] args);

    /// <summary>
    /// 如果逻辑值参数中有一个为 TRUE，则返回 TRUE
    /// 对应 WorksheetFunction.Or 方法
    /// </summary>
    /// <param name="arg1">第一个逻辑值</param>
    /// <param name="args">其他逻辑值</param>
    /// <returns>逻辑或结果</returns>
    bool? Or(object arg1, params object[] args);
    #endregion

    #region 金融函数   
    /// <summary>
    /// 返回某项投资的利率
    /// 对应 WorksheetFunction.Rate 方法
    /// </summary>
    /// <param name="nper">总投资期</param>
    /// <param name="pmt">各期应付给付额</param>
    /// <param name="pv">现值</param>
    /// <param name="fv">未来值</param>
    /// <param name="type">数字 0 或 1，用以指定各期的付款时间是在期初还是期末</param>
    /// <param name="guess">预期利率</param>
    /// <returns>利率</returns>
    double? Rate(double nper, double pmt, double pv, object? fv = null, object? type = null, object? guess = null);

    /// <summary>
    /// 返回某项投资的期数
    /// 对应 WorksheetFunction.NPer 方法
    /// </summary>
    /// <param name="rate">各期利率</param>
    /// <param name="pmt">各期应付给付额</param>
    /// <param name="pv">现值</param>
    /// <param name="fv">未来值</param>
    /// <param name="type">数字 0 或 1，用以指定各期的付款时间是在期初还是期末</param>
    /// <returns>期数</returns>
    double? NPer(double rate, double pmt, double pv, object? fv = null, object? type = null);

    #endregion

}

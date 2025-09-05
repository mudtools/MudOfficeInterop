//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel WorksheetFunction 对象的二次封装实现类
/// 实现 IExcelWorksheetFunction 接口
/// </summary>
internal class ExcelWorksheetFunction : IExcelWorksheetFunction
{
    private MsExcel.WorksheetFunction _worksheetFunction;
    private bool _disposedValue = false;

    internal ExcelWorksheetFunction(MsExcel.WorksheetFunction worksheetFunction)
    {
        _worksheetFunction = worksheetFunction ?? throw new ArgumentNullException(nameof(worksheetFunction));
    }

    #region 基础属性
    public object Parent => _worksheetFunction.Parent;

    public IExcelApplication Application => new ExcelApplication(_worksheetFunction.Application); // 伪代码占位符
    #endregion

    #region 数学和三角函数
    public double Acos(double number) => _worksheetFunction.Acos(number);
    public double Asin(double number) => _worksheetFunction.Asin(number);
    public double Atan2(double xNum, double yNum) => _worksheetFunction.Atan2(xNum, yNum);
    public double Ceiling(double number, double significance) => _worksheetFunction.Ceiling(number, significance);
    public double Combin(double number, double numberChosen) => _worksheetFunction.Combin(number, numberChosen);
    public double Cosh(double number) => _worksheetFunction.Cosh(number);
    public double Degrees(double angle) => _worksheetFunction.Degrees(angle);
    public double Floor(double number, double significance) => _worksheetFunction.Floor(number, significance);
    public double Fact(double number) => _worksheetFunction.Fact(number);
    public double FactDouble(double number) => _worksheetFunction.FactDouble(number);
    public double Ln(double number) => _worksheetFunction.Ln(number);
    public double Log(double number, object baseNumber = null) => _worksheetFunction.Log(number, baseNumber ?? Type.Missing);
    public double Log10(double number) => _worksheetFunction.Log10(number);
    public double Round(double number, int numDigits) => _worksheetFunction.Round(number, numDigits);
    public double RoundUp(double number, int numDigits) => _worksheetFunction.RoundUp(number, numDigits);
    public double RoundDown(double number, int numDigits) => _worksheetFunction.RoundDown(number, numDigits);
    public double Sinh(double number) => _worksheetFunction.Sinh(number);
    public double Tanh(double number) => _worksheetFunction.Tanh(number);
    public double Radians(double angle) => _worksheetFunction.Radians(angle);
    public double Power(double number, double power) => _worksheetFunction.Power(number, power);
    #endregion

    #region 统计函数
    public double Average(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Average(allArgs);
    }
    public double Count(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Count(allArgs);
    }
    public double CountA(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.CountA(allArgs);
    }
    public double Large(object array, double k) => _worksheetFunction.Large(array, k);
    public double Small(object array, double k) => _worksheetFunction.Small(array, k);
    public double Max(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Max(allArgs);
    }
    public double Min(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Min(allArgs);
    }
    public double Product(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Product(allArgs);
    }
    public double Sum(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Sum(allArgs);
    }

    public double StDev(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.StDev(allArgs);
    }
    public double StDevP(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.StDevP(allArgs);
    }
    public double Var(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.Var(allArgs);
    }
    public double VarP(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1 is ExcelRange range ? range.InternalRange : arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return (double)_worksheetFunction.VarP(allArgs);
    }
    #endregion

    #region 文本函数
    public string Fixed(double number, object decimals = null, object noCommas = null) => _worksheetFunction.Fixed(number, decimals ?? Type.Missing, noCommas ?? Type.Missing);
    public string Trim(string text) => _worksheetFunction.Trim(text);
    public double Find(string findText, string withinText, object startNum = null) => _worksheetFunction.Find(findText, withinText, startNum ?? Type.Missing);
    public string Replace(string oldText, double startNum, double numChars, string newText) => _worksheetFunction.Replace(oldText, startNum, numChars, newText);
    public string Rept(string text, double numberTimes) => _worksheetFunction.Rept(text, numberTimes);
    public string Substitute(string text, string oldText, string newText, object instanceNum = null) => _worksheetFunction.Substitute(text, oldText, newText, instanceNum ?? Type.Missing);
    #endregion

    #region 日期和时间函数
    public double Days360(object startDate, object endDate, object method = null) => _worksheetFunction.Days360(startDate, endDate, method ?? Type.Missing);

    public double Weekday(object serialNumber, object returnType = null) => _worksheetFunction.Weekday(serialNumber, returnType ?? Type.Missing);
    public double WeekNum(object serialNumber, object returnType = null) => _worksheetFunction.WeekNum(serialNumber, returnType ?? Type.Missing);
    #endregion

    #region 查找和引用函数
    public object Index(object array, double rowNum, object columnNum = null) => _worksheetFunction.Index(array, rowNum, columnNum ?? Type.Missing);
    public object HLookup(object lookupValue, object tableArray, double rowIndexNum, object rangeLookup = null) => _worksheetFunction.HLookup(lookupValue, tableArray, rowIndexNum, rangeLookup ?? Type.Missing);
    public object VLookup(object lookupValue, object tableArray, double columnIndexNum, object rangeLookup = null) => _worksheetFunction.VLookup(lookupValue, tableArray, columnIndexNum, rangeLookup ?? Type.Missing);

    #endregion

    #region 工程函数
    public string Bin2Dec(string number) => _worksheetFunction.Bin2Dec(number);
    public string Bin2Hex(string number, object places = null) => _worksheetFunction.Bin2Hex(number, places ?? Type.Missing);
    public string Bin2Oct(string number, object places = null) => _worksheetFunction.Bin2Oct(number, places ?? Type.Missing);
    public string Dec2Bin(double number, object places = null) => _worksheetFunction.Dec2Bin(number, places ?? Type.Missing);
    public string Dec2Hex(double number, object places = null) => _worksheetFunction.Dec2Hex(number, places ?? Type.Missing);
    public string Dec2Oct(double number, object places = null) => _worksheetFunction.Dec2Oct(number, places ?? Type.Missing);
    public string Hex2Bin(string number, object places = null) => _worksheetFunction.Hex2Bin(number, places ?? Type.Missing);
    public string Hex2Dec(string number) => _worksheetFunction.Hex2Dec(number);
    public string Hex2Oct(string number, object places = null) => _worksheetFunction.Hex2Oct(number, places ?? Type.Missing);
    public string Oct2Bin(string number, object places = null) => _worksheetFunction.Oct2Bin(number, places ?? Type.Missing);
    public string Oct2Dec(string number) => _worksheetFunction.Oct2Dec(number);
    public string Oct2Hex(string number, object places = null) => _worksheetFunction.Oct2Hex(number, places ?? Type.Missing);
    #endregion

    #region 信息函数
    public bool IsNA(object value) => _worksheetFunction.IsNA(value);
    public bool IsErr(object value) => _worksheetFunction.IsErr(value);
    public bool IsLogical(object value) => _worksheetFunction.IsLogical(value);
    public bool IsNumber(object value) => _worksheetFunction.IsNumber(value);
    public bool IsText(object value) => _worksheetFunction.IsText(value);
    public bool IsEven(object number) => _worksheetFunction.IsEven(number);
    public bool IsOdd(object number) => _worksheetFunction.IsOdd(number);
    public bool IsError(object value) => _worksheetFunction.IsError(value);
    public bool IsNonText(object value) => _worksheetFunction.IsNonText(value);
    #endregion

    #region 逻辑函数、
    public object Transpose(object arg1)
    {
        arg1 = arg1 is ExcelRange range ? range.InternalRange : arg1;
        return _worksheetFunction.Transpose(arg1);
    }

    public bool And(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return _worksheetFunction.And(allArgs);
    }
    public bool Or(object arg1, params object[] args)
    {
        object[] allArgs = new object[args.Length + 1];
        allArgs[0] = arg1;
        Array.Copy(args, 0, allArgs, 1, args.Length);
        return _worksheetFunction.Or(allArgs);
    }
    #endregion

    #region 金融函数
    public double Rate(double nper, double pmt, double pv, object fv = null, object type = null, object guess = null) => _worksheetFunction.Rate(nper, pmt, pv, fv ?? Type.Missing, type ?? Type.Missing, guess ?? Type.Missing);
    public double NPer(double rate, double pmt, double pv, object fv = null, object type = null) => _worksheetFunction.NPer(rate, pmt, pv, fv ?? Type.Missing, type ?? Type.Missing);

    public double PVFactor(double rate, double nper) => 1 / Math.Pow(1 + rate, nper); // Simple implementation
    public double PVAnnuity(double rate, double nper, double pmt, object type = null)
    {
        double effectiveRate = rate;
        double effectiveNPer = nper;
        double effectivePmt = pmt;
        double effectiveType = type == null ? 0 : Convert.ToDouble(type);

        if (effectiveRate == 0)
        {
            return -effectivePmt * effectiveNPer;
        }
        else
        {
            return -effectivePmt * (1 - Math.Pow(1 + effectiveRate, -effectiveNPer)) / effectiveRate * (1 + effectiveRate * effectiveType);
        }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放托管状态(托管对象)
            }

            // 释放未托管的资源(未托管的对象)并在以下内容中重写终结器
            if (_worksheetFunction != null)
            {
                try
                {
                    while (Marshal.ReleaseComObject(_worksheetFunction) > 0) { }
                }
                catch
                {
                    // 忽略释放过程中可能发生的异常
                }
                _worksheetFunction = null;
            }

            _disposedValue = true;
        }
    }

    ~ExcelWorksheetFunction()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}

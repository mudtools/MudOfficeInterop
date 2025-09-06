//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word应用程序中使用的键盘按键枚举
/// </summary>
public enum WdKey
{
    /// <summary>
/// 无按键
/// </summary>
    wdNoKey = 255,
    /// <summary>
/// Shift键
/// </summary>
    wdKeyShift = 256,
    /// <summary>
/// Control键
/// </summary>
    wdKeyControl = 512,
    /// <summary>
/// Command键（Mac系统）
/// </summary>
    wdKeyCommand = 512,
    /// <summary>
/// Alt键
/// </summary>
    wdKeyAlt = 1024,
    /// <summary>
/// Option键（Mac系统）
/// </summary>
    wdKeyOption = 1024,
    /// <summary>
/// 字母A键
/// </summary>
    wdKeyA = 65,
    /// <summary>
/// 字母B键
/// </summary>
    wdKeyB = 66,
    /// <summary>
/// 字母C键
/// </summary>
    wdKeyC = 67,
    /// <summary>
/// 字母D键
/// </summary>
    wdKeyD = 68,
    /// <summary>
/// 字母E键
/// </summary>
    wdKeyE = 69,
    /// <summary>
/// 字母F键
/// </summary>
    wdKeyF = 70,
    /// <summary>
/// 字母G键
/// </summary>
    wdKeyG = 71,
    /// <summary>
/// 字母H键
/// </summary>
    wdKeyH = 72,
    /// <summary>
/// 字母I键
/// </summary>
    wdKeyI = 73,
    /// <summary>
/// 字母J键
/// </summary>
    wdKeyJ = 74,
    /// <summary>
/// 字母K键
/// </summary>
    wdKeyK = 75,
    /// <summary>
/// 字母L键
/// </summary>
    wdKeyL = 76,
    /// <summary>
/// 字母M键
/// </summary>
    wdKeyM = 77,
    /// <summary>
/// 字母N键
/// </summary>
    wdKeyN = 78,
    /// <summary>
/// 字母O键
/// </summary>
    wdKeyO = 79,
    /// <summary>
/// 字母P键
/// </summary>
    wdKeyP = 80,
    /// <summary>
/// 字母Q键
/// </summary>
    wdKeyQ = 81,
    /// <summary>
/// 字母R键
/// </summary>
    wdKeyR = 82,
    /// <summary>
/// 字母S键
/// </summary>
    wdKeyS = 83,
    /// <summary>
/// 字母T键
/// </summary>
    wdKeyT = 84,
    /// <summary>
/// 字母U键
/// </summary>
    wdKeyU = 85,
    /// <summary>
/// 字母V键
/// </summary>
    wdKeyV = 86,
    /// <summary>
/// 字母W键
/// </summary>
    wdKeyW = 87,
    /// <summary>
/// 字母X键
/// </summary>
    wdKeyX = 88,
    /// <summary>
/// 字母Y键
/// </summary>
    wdKeyY = 89,
    /// <summary>
/// 字母Z键
/// </summary>
    wdKeyZ = 90,
    /// <summary>
/// 数字0键
/// </summary>
    wdKey0 = 48,
    /// <summary>
/// 数字1键
/// </summary>
    wdKey1 = 49,
    /// <summary>
/// 数字2键
/// </summary>
    wdKey2 = 50,
    /// <summary>
/// 数字3键
/// </summary>
    wdKey3 = 51,
    /// <summary>
/// 数字4键
/// </summary>
    wdKey4 = 52,
    /// <summary>
/// 数字5键
/// </summary>
    wdKey5 = 53,
    /// <summary>
/// 数字6键
/// </summary>
    wdKey6 = 54,
    /// <summary>
/// 数字7键
/// </summary>
    wdKey7 = 55,
    /// <summary>
/// 数字8键
/// </summary>
    wdKey8 = 56,
    /// <summary>
/// 数字9键
/// </summary>
    wdKey9 = 57,
    /// <summary>
/// 退格键
/// </summary>
    wdKeyBackspace = 8,
    /// <summary>
/// Tab键
/// </summary>
    wdKeyTab = 9,
    /// <summary>
/// 数字键盘5特殊键
/// </summary>
    wdKeyNumeric5Special = 12,
    /// <summary>
/// 回车键
/// </summary>
    wdKeyReturn = 13,
    /// <summary>
/// 暂停键
/// </summary>
    wdKeyPause = 19,
    /// <summary>
/// Esc键
/// </summary>
    wdKeyEsc = 27,
    /// <summary>
/// 空格键
/// </summary>
    wdKeySpacebar = 32,
    /// <summary>
/// Page Up键
/// </summary>
    wdKeyPageUp = 33,
    /// <summary>
/// Page Down键
/// </summary>
    wdKeyPageDown = 34,
    /// <summary>
/// End键
/// </summary>
    wdKeyEnd = 35,
    /// <summary>
/// Home键
/// </summary>
    wdKeyHome = 36,
    /// <summary>
/// Insert键
/// </summary>
    wdKeyInsert = 45,
    /// <summary>
/// Delete键
/// </summary>
    wdKeyDelete = 46,
    /// <summary>
/// 数字键盘0键
/// </summary>
    wdKeyNumeric0 = 96,
    /// <summary>
/// 数字键盘1键
/// </summary>
    wdKeyNumeric1 = 97,
    /// <summary>
/// 数字键盘2键
/// </summary>
    wdKeyNumeric2 = 98,
    /// <summary>
/// 数字键盘3键
/// </summary>
    wdKeyNumeric3 = 99,
    /// <summary>
/// 数字键盘4键
/// </summary>
    wdKeyNumeric4 = 100,
    /// <summary>
/// 数字键盘5键
/// </summary>
    wdKeyNumeric5 = 101,
    /// <summary>
/// 数字键盘6键
/// </summary>
    wdKeyNumeric6 = 102,
    /// <summary>
/// 数字键盘7键
/// </summary>
    wdKeyNumeric7 = 103,
    /// <summary>
/// 数字键盘8键
/// </summary>
    wdKeyNumeric8 = 104,
    /// <summary>
/// 数字键盘9键
/// </summary>
    wdKeyNumeric9 = 105,
    /// <summary>
/// 数字键盘乘法键(*)
/// </summary>
    wdKeyNumericMultiply = 106,
    /// <summary>
/// 数字键盘加法键(+)
/// </summary>
    wdKeyNumericAdd = 107,
    /// <summary>
/// 数字键盘减法键(-)
/// </summary>
    wdKeyNumericSubtract = 109,
    /// <summary>
/// 数字键盘小数点键(.)
/// </summary>
    wdKeyNumericDecimal = 110,
    /// <summary>
/// 数字键盘除法键(/)
/// </summary>
    wdKeyNumericDivide = 111,
    /// <summary>
/// F1功能键
/// </summary>
    wdKeyF1 = 112,
    /// <summary>
/// F2功能键
/// </summary>
    wdKeyF2 = 113,
    /// <summary>
/// F3功能键
/// </summary>
    wdKeyF3 = 114,
    /// <summary>
/// F4功能键
/// </summary>
    wdKeyF4 = 115,
    /// <summary>
/// F5功能键
/// </summary>
    wdKeyF5 = 116,
    /// <summary>
/// F6功能键
/// </summary>
    wdKeyF6 = 117,
    /// <summary>
/// F7功能键
/// </summary>
    wdKeyF7 = 118,
    /// <summary>
/// F8功能键
/// </summary>
    wdKeyF8 = 119,
    /// <summary>
/// F9功能键
/// </summary>
    wdKeyF9 = 120,
    /// <summary>
/// F10功能键
/// </summary>
    wdKeyF10 = 121,
    /// <summary>
/// F11功能键
/// </summary>
    wdKeyF11 = 122,
    /// <summary>
/// F12功能键
/// </summary>
    wdKeyF12 = 123,
    /// <summary>
/// F13功能键
/// </summary>
    wdKeyF13 = 124,
    /// <summary>
/// F14功能键
/// </summary>
    wdKeyF14 = 125,
    /// <summary>
/// F15功能键
/// </summary>
    wdKeyF15 = 126,
    /// <summary>
/// F16功能键
/// </summary>
    wdKeyF16 = 127,
    /// <summary>
/// Scroll Lock键
/// </summary>
    wdKeyScrollLock = 145,
    /// <summary>
/// 分号键(;)
/// </summary>
    wdKeySemiColon = 186,
    /// <summary>
/// 等号键(=)
/// </summary>
    wdKeyEquals = 187,
    /// <summary>
/// 逗号键(,)
/// </summary>
    wdKeyComma = 188,
    /// <summary>
/// 连字符键(-)
/// </summary>
    wdKeyHyphen = 189,
    /// <summary>
/// 句号键(.)
/// </summary>
    wdKeyPeriod = 190,
    /// <summary>
/// 斜杠键(/)
/// </summary>
    wdKeySlash = 191,
    /// <summary>
/// 反引号键(`)
/// </summary>
    wdKeyBackSingleQuote = 192,
    /// <summary>
/// 左方括号键([)
/// </summary>
    wdKeyOpenSquareBrace = 219,
    /// <summary>
/// 反斜杠键(\)
/// </summary>
    wdKeyBackSlash = 220,
    /// <summary>
/// 右方括号键(])
/// </summary>
    wdKeyCloseSquareBrace = 221,
    /// <summary>
/// 单引号键(')
/// </summary>
    wdKeySingleQuote = 222
}
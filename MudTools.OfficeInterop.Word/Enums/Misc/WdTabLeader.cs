namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在制表符和文本之间显示的前导字符类型
/// </summary>
public enum WdTabLeader
{
    /// <summary>
    /// 空格前导字符
    /// </summary>
    wdTabLeaderSpaces,
    /// <summary>
    /// 点前导字符
    /// </summary>
    wdTabLeaderDots,
    /// <summary>
    /// 破折号前导字符
    /// </summary>
    wdTabLeaderDashes,
    /// <summary>
    /// 行前导字符
    /// </summary>
    wdTabLeaderLines,
    /// <summary>
    /// 粗线前导字符
    /// </summary>
    wdTabLeaderHeavy,
    /// <summary>
    /// 中点前导字符
    /// </summary>
    wdTabLeaderMiddleDot
}
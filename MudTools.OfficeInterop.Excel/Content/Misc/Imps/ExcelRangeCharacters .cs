namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelRangeCharacters : ExcelCharacters, IExcelRangeCharacters
{
    private MsExcel.Range? _range;
    public ExcelRangeCharacters(MsExcel.Range range, MsExcel.Characters characters) : base(characters)
    {
        _range = range ?? throw new ArgumentNullException(nameof(range));
    }

    public IExcelCharacters this[int start, int length]
    {
        get
        {
            if (_range == null)
                throw new ObjectDisposedException(nameof(ExcelRangeCharacters));
            if (start < 1 || length < 1 || start + length > Count)
                throw new ArgumentOutOfRangeException(nameof(start), "Start and length must be greater than zero.");

            return new ExcelRangeCharacters(_range, _range.Characters[start, length]);
        }
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing && _range != null)
        {
            Marshal.ReleaseComObject(_range);
            _range = null;
        }
    }
}
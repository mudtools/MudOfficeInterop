
namespace MudTools.OfficeInterop.Word.Imps;

partial class WordDocument
{

    public int PageCount
    {
        get
        {
            if (_document == null) return 0;
            return (int)_document.Range().Information[MsWord.WdInformation.wdNumberOfPagesInDocument];
        }
    }

    public string Title
    {
        get
        {
            return GetBuiltInDocumentProperty("Title");
        }
        set
        {
            SetBuiltInDocumentProperty("Title", value);
        }
    }

    public string Author
    {
        get
        {
            return GetBuiltInDocumentProperty("Author");
        }
        set
        {
            SetBuiltInDocumentProperty("Author", value);
        }
    }

    public string Subject
    {
        get
        {
            return GetBuiltInDocumentProperty("Subject");
        }
        set
        {
            SetBuiltInDocumentProperty("Subject", value);
        }
    }

    public string Description
    {
        get
        {
            return GetBuiltInDocumentProperty("Comments");
        }
        set
        {
            SetBuiltInDocumentProperty("Comments", value);
        }
    }

    public string Keywords
    {
        get
        {
            return GetBuiltInDocumentProperty("Keywords");
        }
        set
        {
            SetBuiltInDocumentProperty("Keywords", value);
        }
    }

    public string Company
    {
        get
        {
            return GetBuiltInDocumentProperty("Company");
        }
        set
        {
            SetBuiltInDocumentProperty("Company", value);
        }
    }


    private string GetBuiltInDocumentProperty(string propertyName)
    {
        try
        {
            if (_document == null)
                return string.Empty;
            if (BuiltInDocumentProperties == null)
                return string.Empty;

            var value = BuiltInDocumentProperties[propertyName]?.Value;
            return value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void SetBuiltInDocumentProperty(string propertyName, string value)
    {
        try
        {
            if (_document == null)
                return;
            if (BuiltInDocumentProperties == null)
                return;
            var property = BuiltInDocumentProperties[propertyName];
            if (property != null)
                property.Value = value;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to set document property '{propertyName}'.", ex);
        }
    }
}
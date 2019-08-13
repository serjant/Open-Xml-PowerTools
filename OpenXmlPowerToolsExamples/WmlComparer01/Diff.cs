using System;

public class Diff
{
    private string _mergedContent;
    public string mergedContent
    {
        get
        {
            return _mergedContent;
        }

        set
        {
            _mergedContent = value;
        }
    }
    public long mergeChangesCounter { get; set; }
}


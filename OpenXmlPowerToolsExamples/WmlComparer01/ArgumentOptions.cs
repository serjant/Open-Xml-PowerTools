using System;
using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;

public class ArgumentOptions
{
    [Option("old-content-file", Required = false, HelpText = "DOCX file path to the old version of the content to diff")]
    public string OldContentFile { get; set; }

    [Option('o', "old-content", Required = false, HelpText = "Old version of the DOCX content to diff")]
    public string OldContent { get; set; }

    [Option("new-content-file", Required = false, HelpText = "DOCX file path to the new version of the content to diff")]
    public string NewContentFile { get; set; }

    [Option('n', "new-content", Required = false, HelpText = "New version of the DOCX content to diff")]
    public string NewContent { get; set; }

    [Option("previous-content-file", Required = false, HelpText = "DOCX file path to the previous version of the content to diff")]
    public string PreviousContentFile { get; set; }

    [Option('p', "previous-content", Required = false, HelpText = "Previous version of the DOCX content to diff")]
    public string PreviousContent { get; set; }

    [Option('u', "username", Required = false, HelpText = "Username of the user who applies the diff")]
    public string Username { get; set; }

    [Usage(ApplicationAlias = "mono WmlComparer.exe")]
    public static IEnumerable<Example> Examples
    {
        get
        {
            return new List<Example>() {
                new Example(
                    "Input file pathes",
                    UnParserSettings.WithUseEqualTokenOnly(),
                    new ArgumentOptions { OldContentFile = "file.bin", NewContentFile = "file.bin", PreviousContentFile = "file.bin", Username = "content.string" }
                ),
                new Example(
                    "Input OOXML contents",
                    UnParserSettings.WithUseEqualTokenOnly(),
                    new ArgumentOptions { OldContent = "content.string", NewContent = "content.string", PreviousContent = "content.string", Username = "content.string" }
                ),
            };
        }
    }
}

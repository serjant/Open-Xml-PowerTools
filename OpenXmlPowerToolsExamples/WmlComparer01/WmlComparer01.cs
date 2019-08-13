// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Text;
using CommandLine;
using CommandLine.Text;
using OpenXmlPowerTools;
using Newtonsoft.Json;

class WmlComparer01
{
    public static bool IsDebug = false;
    public static string DownloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

    private static Diff CompareDocuments(WmlDocument leftPrevDocument, WmlDocument leftDocument, WmlDocument rightDocument, WmlComparerSettings settings)
    {
        WmlDocument result = WmlComparer.Compare(
            leftPrevDocument,
                leftDocument,
                rightDocument,
                settings
            );

        var revisions = WmlComparer.GetRevisions(result, settings);

        if (IsDebug)
            result.SaveAs(Path.Combine(DownloadsFolder, "Compared.docx"));

        Diff diff = new Diff();
        diff.mergedContent = result.toOOXML();
        diff.mergeChangesCounter = revisions.Count;

        return diff;
    }

    private static Diff CompareDocuments(WmlDocument leftDocument, WmlDocument rightDocument, WmlComparerSettings settings)
    {
        WmlDocument result = WmlComparer.Compare(
                leftDocument,
                rightDocument,
                settings
            );

        if (IsDebug)
            result.SaveAs(Path.Combine(DownloadsFolder, "Compared.docx"));

        var revisions = WmlComparer.GetRevisions(result, settings);

        Diff diff = new Diff();
        diff.mergedContent = result.toOOXML();
        diff.mergeChangesCounter = revisions.Count;

        return diff;
    }

    static void Main(string[] args)
    {
        ParserResult<ArgumentOptions> parserResult = CommandLine.Parser.Default.ParseArguments<ArgumentOptions>(args);

        parserResult.WithParsed<ArgumentOptions>(opts =>
        {
            string oldContentFile = null;
            string newContentFile = null;

            if (!IsDebug)
            {
                if (opts.NewContentFile == null)
                {
                    ThrowError("New content file should not be an empty", parserResult);
                }
                if (opts.OldContentFile == null)
                {
                    ThrowError("Old content file should not be an empty", parserResult);
                }

                oldContentFile = opts.OldContentFile;
                newContentFile = opts.NewContentFile;
            }
            else
            {
                oldContentFile = Path.Combine(DownloadsFolder, "2.docx");
                newContentFile = Path.Combine(DownloadsFolder, "3.docx");
                opts.PreviousContentFile = Path.Combine(DownloadsFolder, "1.docx");
            }

            WmlDocument oldDocument = opts.OldContent != null ? new WmlDocument("old.docx", Encoding.Unicode.GetBytes(opts.OldContent)) : new WmlDocument(oldContentFile);
            WmlDocument newDocument = opts.NewContent != null ? new WmlDocument("new.docx", Encoding.Unicode.GetBytes(opts.NewContent)) : new WmlDocument(newContentFile);
            WmlDocument prevDocument = null;
            if (opts.PreviousContentFile != null)
            {
                prevDocument = new WmlDocument(opts.PreviousContentFile);
            }
            if (opts.PreviousContent != null && opts.PreviousContentFile == null)
            {
                prevDocument = new WmlDocument("prev.docx", Encoding.Unicode.GetBytes(opts.PreviousContent));
            }

            WmlComparerSettings settings = new WmlComparerSettings();
            if (opts.Username != null)
            {
                settings.AuthorForRevisions = opts.Username;
            }
            else
            {
                settings.AuthorForRevisions = "Microsoft";
            }

            Diff diff = null;

            if (prevDocument == null)
            {
                diff = CompareDocuments(
                    oldDocument,
                    newDocument,
                    settings
                    );
            }
            else
            {
                diff = CompareDocuments(
                    prevDocument,
                    oldDocument,
                    newDocument,
                    settings
                    );
            }
            var json = JsonConvert.SerializeObject(diff);
            Console.WriteLine(json);
        }).WithNotParsed<ArgumentOptions>((errs) =>
        {
            string exceptionString = null;
            foreach (var error in errs)
            {
                if (exceptionString == null)
                {
                    exceptionString = error.ToString();
                }
                else
                {
                    exceptionString = String.Format("{0}\n{1}", exceptionString, error.ToString());
                }
            }
            Exception exception = new Exception(exceptionString);
            throw exception;
        });
    }

    public static void ThrowError(string error, ParserResult<ArgumentOptions> parserResult)
    {
        string helpText = HelpText.RenderUsageText(parserResult);
        Console.WriteLine(helpText);

        Exception exception = new Exception(error);
        throw exception;
    }
}
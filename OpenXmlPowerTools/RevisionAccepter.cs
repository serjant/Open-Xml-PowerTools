// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class RevisionAccepter
    {
        public static WmlDocument ReverseDeletedRevisions(WmlDocument document)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    ReverseDeletedRevisions(wDoc);
                }
                return new WmlDocument(document.FileName, ms.ToArray());
            }
        }

        public static WmlDocument AcceptRevisions(WmlDocument document)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    AcceptRevisions(wDoc);
                }
                return new WmlDocument(document.FileName, ms.ToArray());
            }
        }

        public static WmlDocument RejectRevisions(WmlDocument document)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    RejectRevisions(wDoc);
                }
                return new WmlDocument(document.FileName, ms.ToArray());
            }
        }

        public static void ReverseDeletedRevisions(WordprocessingDocument doc)
        {
            RevisionProcessor.ReverseDeletedRevisions(doc);
        }

        public static void RejectRevisions(WordprocessingDocument doc)
        {
            RevisionProcessor.RejectRevisions(doc);
        }

        public static void AcceptRevisions(WordprocessingDocument doc)
        {
            RevisionProcessor.AcceptRevisions(doc);
        }

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            return RevisionProcessor.PartHasTrackedRevisions(part);
        }

        public static bool HasTrackedRevisions(WmlDocument document)
        {
            return RevisionProcessor.HasTrackedRevisions(document);
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            return RevisionProcessor.HasTrackedRevisions(doc);
        }
    }
}

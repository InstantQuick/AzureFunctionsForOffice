using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using OXML;

namespace AzureFunctionsForOffice.WordMerge
{
    public class Merger
    {
        public static byte[] MergeDocumentWithContent(byte[] documentTemplate, Dictionary<string, string> fieldData, Dictionary<string, List<Dictionary<string, string>>> repeatingSectionData)
        {
            var fileMemStream = new MemoryStream();
            fileMemStream.Write(documentTemplate, 0, documentTemplate.Length);
            using (var doc = WordprocessingDocument.Open(fileMemStream, true))
            {
                MainDocumentPart mainDocumentPart = doc.MainDocumentPart;
                Document document = mainDocumentPart.Document;

                ProcessFieldData(fieldData, document);
                ProcessRepeatingSections(repeatingSectionData, document);
                document.Save();
            }
            return fileMemStream.ToArray();
        }

        private static void ProcessFieldData(Dictionary<string, string> fieldData, Document document)
        {
            if (fieldData == null) return;
            var fieldKeys = fieldData.Keys.ToList();
            var matchingElements = GetPlaceHolderElements(document, fieldKeys);

            foreach (var key in fieldKeys)
            {
                SetContentOfContentControl(matchingElements[key] as SdtElement, fieldData[key]);
            }
        }

        private static void ProcessRepeatingSections(Dictionary<string, List<Dictionary<string, string>>> repeatingSectionData, Document document)
        {
            if (repeatingSectionData == null) return;
            var repeatingSectionKeys = repeatingSectionData.Keys.ToList();
            var repeatingSectionElements = GetPlaceHolderElements(document, repeatingSectionKeys);

            foreach (var key in repeatingSectionKeys)
            {
                FillRepeatingSection(repeatingSectionElements[key], repeatingSectionData[key]);
            }
        }

        static void FillRepeatingSection(OpenXmlElement element, List<Dictionary<string, string>> data)
        {
            if (element == null) { return; }

            GetTagValue(element as SdtElement, out _, out _);

            foreach (var item in data)
            {
                var clonedSdtElement = element.InsertBeforeSelf(element.CloneNode(true) as SdtElement);

                var fieldPlaceHolderNames = item.Keys.ToList();

                Dictionary<string, OpenXmlElement> elements = new Dictionary<string, OpenXmlElement>();
                LookForPlaceHoldersInChildren(clonedSdtElement, fieldPlaceHolderNames, elements);

                foreach (string key in elements.Keys)
                {
                    SetContentOfContentControl(elements[key] as SdtElement, item[key]);
                }
            }

            element.Remove();
        }

        static Dictionary<string, OpenXmlElement> GetPlaceHolderElements(Document document, List<string> placeHolderNames)
        {
            var elements = new Dictionary<string, OpenXmlElement>();
            LookForPlaceHolders(document, placeHolderNames, elements);
            return elements;
        }

        static void LookForPlaceHolders(OpenXmlElement element, List<string> placeHolderNames, Dictionary<string, OpenXmlElement> placeHolderElements)
        {
            if (IsContentControl(element))
            {
                SdtElement sdtElement = element as SdtElement;
                GetTagValue(sdtElement, out var templateTagPart, out _);

                if (placeHolderNames.Contains(templateTagPart))
                {
                    placeHolderElements[templateTagPart] = element;
                }
            }
            LookForPlaceHoldersInChildren(element, placeHolderNames, placeHolderElements);
        }

        static void LookForPlaceHoldersInChildren(OpenXmlElement element, List<string> placeHolderNames, Dictionary<string, OpenXmlElement> placeHolderElements)
        {
            if (element is OpenXmlCompositeElement && element.HasChildren)
            {
                List<OpenXmlElement> elements = element.Elements().ToList();

                foreach (var childElement in elements)
                {
                    if (childElement is OpenXmlCompositeElement)
                    {
                        LookForPlaceHolders(childElement, placeHolderNames, placeHolderElements);
                    }
                }
            }
        }

        static bool IsContentControl(OpenXmlElement element)
        {
            return element != null && (element is SdtBlock || element is SdtRun || element is SdtRow || element is SdtCell);
        }

        static void GetTagValue(SdtElement element, out string templateTagPart, out string tagGuidPart)
        {
            templateTagPart = string.Empty;
            tagGuidPart = string.Empty;
            Tag tag = GetTag(element);

            string fullTag = (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;

            if (!string.IsNullOrEmpty(fullTag))
            {
                string[] tagParts = fullTag.Split(':');

                if (tagParts.Length == 2)
                {
                    templateTagPart = tagParts[0];
                    tagGuidPart = tagParts[1];
                }
                else if (tagParts.Length == 1)
                {
                    templateTagPart = tagParts[0];
                }
            }
        }
        static Tag GetTag(SdtElement element)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            return element.SdtProperties.Elements<Tag>().FirstOrDefault();
        }

        static readonly OpenXmlHelper openXmlHelper = new OpenXmlHelper();
        static void SetContentOfContentControl(SdtElement element, string content)
        {
            openXmlHelper.SetContentOfContentControl(element, content);
        }
    }
}

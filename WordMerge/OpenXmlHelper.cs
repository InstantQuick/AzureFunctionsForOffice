namespace OXML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// Helper class for OpenXml operations for document generation
    /// </summary>
    public class OpenXmlHelper
    {
        #region Public Methods

        /// <summary>
        /// Gets the SDT content of content control.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public OpenXmlCompositeElement GetSdtContentOfContentControl(SdtElement element)
        {
            SdtRun sdtRunELement = element as SdtRun;
            SdtBlock sdtBlockElement = element as SdtBlock;
            SdtCell sdtCellElement = element as SdtCell;
            SdtRow sdtRowElement = element as SdtRow;

            if (sdtRunELement != null)
            {
                return sdtRunELement.SdtContentRun;
            }
            else if (sdtBlockElement != null)
            {
                return sdtBlockElement.SdtContentBlock;
            }
            else if (sdtCellElement != null)
            {
                return sdtCellElement.SdtContentCell;
            }
            else if (sdtRowElement != null)
            {
                return sdtRowElement.SdtContentRow;
            }

            return null;
        }
        
        /// <summary>
        /// Sets the content of content control.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <param name="content">The content.</param>
        public void SetContentOfContentControl(SdtElement contentControl, string content)
        {
            if (contentControl == null)
            {
                throw new ArgumentNullException(nameof(contentControl));
            }

            content = string.IsNullOrEmpty(content) ? string.Empty : content;
            bool isCombobox = contentControl.SdtProperties.Descendants<SdtContentDropDownList>().FirstOrDefault() != null;

            if (isCombobox)
            {
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                Run run = CreateRun(openXmlCompositeElement, content);
                SetSdtContentKeepingPermissionElements(openXmlCompositeElement, run);
            }
            else
            {
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                contentControl.SdtProperties.RemoveAllChildren<ShowingPlaceholder>();
                List<Run> runs = new List<Run>();

                if (content.Contains(Environment.NewLine))
                {
                    List<string> lines = content.Split(Environment.NewLine.ToCharArray()).ToList();

                    foreach (string line in lines)
                    {
                        Run run = CreateRun(openXmlCompositeElement, line);

                        if (string.IsNullOrEmpty(line))
                        {
                            run.AppendChild(new Break());
                        }

                        runs.Add(run);
                    }
                }
                else
                {
                    runs.Add(CreateRun(openXmlCompositeElement, content));
                }

                if (openXmlCompositeElement is SdtContentCell)
                {
                    AddRunsToSdtContentCell(openXmlCompositeElement as SdtContentCell, runs);
                }
                else if (openXmlCompositeElement is SdtContentBlock)
                {
                    Paragraph para = CreateParagraph(openXmlCompositeElement, runs);
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, para);
                }
                else
                {
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, runs);
                }
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChild">The new child.</param>
        private void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, OpenXmlElement newChild)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            openXmlCompositeElement.AppendChild(newChild);

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChildren">The new children.</param>
        private void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, List<Run> newChildren)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            foreach (var newChild in newChildren)
            {
                openXmlCompositeElement.AppendChild(newChild);
            }

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Adds the runs to SDT content cell.
        /// </summary>
        /// <param name="sdtContentCell">The SDT content cell.</param>
        /// <param name="runs">The runs.</param>
        private void AddRunsToSdtContentCell(SdtContentCell sdtContentCell, List<Run> runs)
        {
            TableCell cell = new TableCell();
            Paragraph para = new Paragraph();
            para.RemoveAllChildren();

            foreach (Run run in runs)
            {
                para.AppendChild(run);
            }

            cell.AppendChild(para);
            SetSdtContentKeepingPermissionElements(sdtContentCell, cell);
        }

        /// <summary>
        /// Creates the paragraph.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="runs">The runs.</param>
        /// <returns></returns>
        private static Paragraph CreateParagraph(OpenXmlCompositeElement openXmlCompositeElement, List<Run> runs)
        {
            ParagraphProperties paragraphProperties = openXmlCompositeElement.Descendants<ParagraphProperties>().FirstOrDefault();
            Paragraph para;

            if (paragraphProperties != null)
            {
                para = new Paragraph(paragraphProperties.CloneNode(true));
                foreach (Run run in runs)
                {
                    para.AppendChild(run);
                }
            }
            else
            {
                para = new Paragraph();
                foreach (Run run in runs)
                {
                    para.AppendChild(run);
                }
            }
            return para;
        }

        /// <summary>
        /// Creates the run.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="content">The content.</param>
        /// <returns></returns>
        private static Run CreateRun(OpenXmlCompositeElement openXmlCompositeElement, string content)
        {
            RunProperties runProperties = openXmlCompositeElement.Descendants<RunProperties>().FirstOrDefault();
            return runProperties != null ? new Run(runProperties.CloneNode(true), new Text(content)) : new Run(new Text(content));
        }

        #endregion
    }
}

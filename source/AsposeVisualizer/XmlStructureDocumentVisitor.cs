// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XmlStructureDocumentVisitor.cs" company="Philipp Dolder">
//   Copyright (c) 2013
//
//   Licensed under the Apache License, Version 2.0 (the "License");
//   you may not use this file except in compliance with the License.
//   You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//   Unless required by applicable law or agreed to in writing, software
//   distributed under the License is distributed on an "AS IS" BASIS,
//   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//   See the License for the specific language governing permissions and
//   limitations under the License.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------
namespace AsposeVisualizer
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;
    using Aspose.Words;
    using Aspose.Words.Fields;

    public class XmlStructureDocumentVisitor : DocumentVisitor
    {
        private const string DocumentVariablePrefix = "docvariable";
        private const string DocumentPropertyPrefix = "docproperty";
        private readonly StringBuilder structureBuilder;
        private bool skipRun;
        private bool skipFieldSeparator;
        private string currentFieldTagName;
        private string currentDocumentProperty;
        private string currentDocumentVariable;

        public XmlStructureDocumentVisitor()
        {
            this.structureBuilder = new StringBuilder();
            this.skipRun = false;
            this.skipFieldSeparator = false;
        }

        public string AsXml
        {
            get { return FormatXml(this.structureBuilder.ToString()); }
        }

        public override VisitorAction VisitShapeStart(Aspose.Words.Drawing.Shape shape)
        {
            this.structureBuilder.AppendFormat("<Shape width=\"{0}\" height=\"{1}\">", shape.Width, shape.Height);
            this.structureBuilder.AppendLine();

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitShapeEnd(Aspose.Words.Drawing.Shape shape)
        {
            this.structureBuilder.AppendLine("</Shape>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableStart(Aspose.Words.Tables.Table table)
        {
            this.structureBuilder.AppendLine("<Table>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableEnd(Aspose.Words.Tables.Table table)
        {
            this.structureBuilder.AppendLine("</Table>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellStart(Aspose.Words.Tables.Cell cell)
        {
            this.structureBuilder.AppendLine("<Cell>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellEnd(Aspose.Words.Tables.Cell cell)
        {
            this.structureBuilder.AppendLine("</Cell>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowStart(Aspose.Words.Tables.Row row)
        {
            this.structureBuilder.AppendLine("<Row>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowEnd(Aspose.Words.Tables.Row row)
        {
            this.structureBuilder.AppendLine("</Row>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDocumentStart(Document doc)
        {
            this.structureBuilder.AppendLine("<Document>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDocumentEnd(Document doc)
        {
            this.structureBuilder.AppendLine("</Document>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionStart(Section section)
        {
            this.structureBuilder
                .AppendFormat("<Section PaperSize='{0}'>", section.PageSetup.PaperSize)
                .AppendLine();

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionEnd(Section section)
        {
            this.structureBuilder.AppendLine("</Section>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBodyStart(Body body)
        {
            this.structureBuilder.AppendLine("<Body>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBodyEnd(Body body)
        {
            this.structureBuilder.AppendLine("</Body>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            this.structureBuilder
                .AppendFormat("<Paragraph StyleIdentifier='{0}' StyleName='{1}'>", paragraph.ParagraphFormat.StyleIdentifier, paragraph.ParagraphFormat.StyleName)
                .AppendLine();

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            this.structureBuilder.AppendLine("</Paragraph>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
        {
            this.structureBuilder
                .AppendFormat("<BookmarkStart Name='{0}' />", bookmarkStart.Name)
                .AppendLine();

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
        {
            this.structureBuilder
                .AppendFormat("<BookmarkEnd Name='{0}' />", bookmarkEnd.Name)
                .AppendLine();

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRun(Run run)
        {
            if (this.skipRun)
            {
                this.skipRun = false;
                this.skipFieldSeparator = true;
            }
            else if (run.Text.Contains(ControlChar.PageBreak))
            {
                this.structureBuilder
                    .AppendLine("<PageBreak />");
            }
            else
            {
                this.structureBuilder
                    .AppendFormat("<Run>{0}</Run>", run.Text)
                    .AppendLine();
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBuildingBlockStart(Aspose.Words.BuildingBlocks.BuildingBlock block)
        {
            this.structureBuilder.AppendLine("<BuildingBlock>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBuildingBlockEnd(Aspose.Words.BuildingBlocks.BuildingBlock block)
        {
            this.structureBuilder.AppendLine("</BuildingBlock>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCommentStart(Comment comment)
        {
            this.structureBuilder.AppendLine("<Comment>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCommentEnd(Comment comment)
        {
            this.structureBuilder.AppendLine("</Comment>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
        {
            this.structureBuilder.AppendLine("<CommentRange>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
        {
            this.structureBuilder.AppendLine("</CommentRange>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCustomXmlMarkupStart(Aspose.Words.Markup.CustomXmlMarkup customXmlMarkup)
        {
            this.structureBuilder.AppendLine("<CustomXmlMarkup>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCustomXmlMarkupEnd(Aspose.Words.Markup.CustomXmlMarkup customXmlMarkup)
        {
            this.structureBuilder.AppendLine("</CustomXmlMarkup>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDrawingML(Aspose.Words.Drawing.DrawingML drawingMl)
        {
            this.structureBuilder.AppendLine("<DrawingML />");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldStart(FieldStart fieldStart)
        {
            if (fieldStart.FieldType == FieldType.FieldDocProperty)
            {
                this.currentDocumentProperty = GetDocumentPropertyFromField(fieldStart);
                this.currentFieldTagName = GetXmlTagForDocumentProperty(this.currentDocumentProperty);
                this.structureBuilder.AppendLine(string.Concat("<", this.currentFieldTagName, "Start Name=\"", this.currentDocumentProperty, "\" />"));
                this.skipRun = true;
            }
            else if (fieldStart.FieldType == FieldType.FieldDocVariable)
            {
                this.currentDocumentVariable = GetDocumentVariableFromField(fieldStart);
                this.currentFieldTagName = "DocumentVariable";
                this.structureBuilder.AppendLine(string.Concat("<", this.currentFieldTagName, "Start Name=\"", this.currentDocumentVariable, "\" />"));
                this.skipRun = true;
            }
            else
            {
                this.structureBuilder.AppendLine("<FieldStart />");
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
        {
            if (!string.IsNullOrEmpty(this.currentDocumentProperty))
            {
                this.structureBuilder.AppendLine(string.Concat("<", this.currentFieldTagName, "End Name=\"", this.currentDocumentProperty, "\"/>"));
                this.currentDocumentProperty = null;
            }
            else if (!string.IsNullOrEmpty(this.currentDocumentVariable))
            {
                this.structureBuilder.AppendLine(string.Concat("<", this.currentFieldTagName, "End Name=\"", this.currentDocumentVariable, "\"/>"));
                this.currentDocumentVariable = null;
            }
            else
            {
                this.structureBuilder.AppendLine("<FieldEnd />");
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
        {
            if (this.skipFieldSeparator)
            {
                this.skipFieldSeparator = false;
            }
            else
            {
                this.structureBuilder.AppendLine("<FieldSeparator />");
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFootnoteStart(Footnote footnote)
        {
            this.structureBuilder.AppendLine("<Footnote>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFootnoteEnd(Footnote footnote)
        {
            this.structureBuilder.AppendLine("</Footnote>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFormField(FormField formField)
        {
            this.structureBuilder.AppendLine("<FormField />");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGlossaryDocumentStart(Aspose.Words.BuildingBlocks.GlossaryDocument glossary)
        {
            this.structureBuilder.AppendLine("<GlossaryDocument>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGlossaryDocumentEnd(Aspose.Words.BuildingBlocks.GlossaryDocument glossary)
        {
            this.structureBuilder.AppendLine("</GlossaryDocument>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGroupShapeStart(Aspose.Words.Drawing.GroupShape groupShape)
        {
            this.structureBuilder.AppendLine("<GroupShape>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGroupShapeEnd(Aspose.Words.Drawing.GroupShape groupShape)
        {
            this.structureBuilder.AppendLine("</GroupShape>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
        {
            this.structureBuilder.AppendLine(headerFooter.IsHeader ? "<Header>" : "<Footer>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
        {
            this.structureBuilder.AppendLine(headerFooter.IsHeader ? "</Header>" : "</Footer>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitOfficeMathStart(Aspose.Words.Math.OfficeMath officeMath)
        {
            this.structureBuilder.AppendLine("<OfficeMath>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitOfficeMathEnd(Aspose.Words.Math.OfficeMath officeMath)
        {
            this.structureBuilder.AppendLine("</OfficeMath>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSmartTagStart(Aspose.Words.Markup.SmartTag smartTag)
        {
            this.structureBuilder.AppendLine("<SmartTag>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSmartTagEnd(Aspose.Words.Markup.SmartTag smartTag)
        {
            this.structureBuilder.AppendLine("</SmartTag>");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSpecialChar(SpecialChar specialChar)
        {
            this.structureBuilder.AppendLine("<SpecialChar />");

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagStart(Aspose.Words.Markup.StructuredDocumentTag sdt)
        {
            this.structureBuilder.AppendLine(string.Concat("<ContentControl Title=\"", sdt.Title, "\" >"));

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagEnd(Aspose.Words.Markup.StructuredDocumentTag sdt)
        {
            this.structureBuilder.AppendLine("</ContentControl>");

            return VisitorAction.Continue;
        }

        private static string GetXmlTagForDocumentProperty(string documentProperty)
        {
            return IsBuiltInDocumentProperty(documentProperty) ? "BuiltInDocumentProperty" : "CustomDocumentProperty";
        }

        private static string GetDocumentPropertyFromField(FieldStart fieldStart)
        {
            var nextSibling = fieldStart.NextSibling as Run;
            var runText = nextSibling.Text;
            int lengthOfDocumentPropertyPrefix = DocumentPropertyPrefix.Length;

            return runText.Substring(lengthOfDocumentPropertyPrefix + 1).Trim();
        }

        private static string GetDocumentVariableFromField(FieldStart fieldStart)
        {
            var nextSibling = fieldStart.NextSibling as Run;
            var runText = nextSibling.Text;
            int lengthOfDocumentVariablePrefix = DocumentVariablePrefix.Length;

            return runText.Substring(lengthOfDocumentVariablePrefix + 1).Trim();
        }

        private static bool IsBuiltInDocumentProperty(string documentPropertyName)
        {
            var builtInDocumentPropertyNames = GetConstantsOf(typeof(BuiltInDocumentPropertyNames));

            return builtInDocumentPropertyNames.Contains(documentPropertyName);
        }

        private static List<string> GetConstantsOf(Type type)
        {
            var fieldInfos = type.GetFields(BindingFlags.Public | BindingFlags.Static);

            return fieldInfos.Select(fieldInfo => fieldInfo.Name).ToList();
        }

        private static string FormatXml(string xml)
        {
            try
            {
                return XElement.Parse(xml).ToString(SaveOptions.None);
            }
            catch (XmlException)
            {
                string message = new StringBuilder()
                    .AppendLine("Could not interpret the document as xml.")
                    .AppendLine("Please raise an issue at https://github.com/philippdolder/AsposeVisualizer to help improve the aspose debugger visualizer.")
                    .AppendLine()
                    .AppendLine("Here is the document as string representation:")
                    .AppendLine()
                    .Append(xml)
                    .ToString();

                return message;
            }
        }
    }
}
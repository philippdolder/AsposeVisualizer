// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XmlStructureNodeVisitor.cs" company="Philipp Dolder">
//   Copyright (c) 2014
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
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;

    public class XmlStructureNodeVisitor : NodeVisitor
    {
        private readonly XmlStructureDisplayOptions displayOptions;
        private readonly StringBuilder builder = new StringBuilder();

        public XmlStructureNodeVisitor(XmlStructureDisplayOptions displayOptions)
        {
            this.displayOptions = displayOptions;
        }

        public string AsXml
        {
            get
            {
                return XElement.Parse(this.builder.ToString()).ToString(SaveOptions.None);
            }
        }

        public override void VisitDocumentStart(DocumentProxy document)
        {
            this.builder.AppendLine("<Document>");
        }

        public override void VisitDocumentEnd(DocumentProxy document)
        {
            this.builder.AppendLine("</Document>");
        }

        public override void VisitSectionStart(SectionProxy section)
        {
            if (this.displayOptions.IncludeFormatting)
            {
                this.builder
                    .AppendFormat(
                        "<Section {0}>",
                        FormatAttributes(
                            new NamedValue("Orientation", section.Format.Orientation),
                            new NamedValue("PaperSize", section.Format.PaperSize)))
                    .AppendLine();
            }
            else
            {
                this.builder.AppendLine("<Section>");
            }
        }

        public override void VisitSectionEnd(SectionProxy section)
        {
            this.builder.AppendLine("</Section>");
        }

        public override void VisitBodyStart(BodyProxy body)
        {
            this.builder.AppendLine("<Body>");
        }

        public override void VisitBodyEnd(BodyProxy body)
        {
            this.builder.AppendLine("</Body>");
        }

        public override void VisitHeaderStart(HeaderProxy header)
        {
            this.builder
                .AppendFormat(
                    "<Header {0}>", 
                    FormatAttributes(new NamedValue("Type", header.Type), new NamedValue("LinkedToPrevious", header.IsLinkedToPrevious.ToString())))
                .AppendLine();
        }

        public override void VisitHeaderEnd(HeaderProxy header)
        {
            this.builder.AppendLine("</Header>");
        }

        public override void VisitFooterStart(FooterProxy footer)
        {
            this.builder
                .AppendFormat(
                    "<Footer {0}>",
                    FormatAttributes(new NamedValue("Type", footer.Type), new NamedValue("LinkedToPrevious", footer.IsLinkedToPrevious.ToString())))
                .AppendLine();
        }

        public override void VisitFooterEnd(FooterProxy footer)
        {
            this.builder.AppendLine("</Footer>");
        }

        public override void VisitParagraphStart(ParagraphProxy paragraph)
        {
            if (this.displayOptions.IncludeFormatting)
            {
                this.builder
                    .AppendFormat(
                        "<Paragraph {0}>",
                        FormatAttributes(
                            new NamedValue("StyleIdentifier", paragraph.Format.StyleIdentifier),
                            new NamedValue("StyleName", paragraph.Format.StyleName)))
                    .AppendLine();
            }
            else
            {
                this.builder.AppendLine("<Paragraph>");
            }
        }

        public override void VisitParagraphEnd(ParagraphProxy paragraph)
        {
            this.builder.AppendLine("</Paragraph>");
        }

        public override void VisitRun(RunProxy run)
        {
            if (this.displayOptions.IncludeFormatting)
            {
                this.builder
                    .AppendFormat(
                        "<Run {0}>{1}</Run>",
                        FormatAttributes(
                            new NamedValue("Language", run.Format.Language.ToString(CultureInfo.InvariantCulture)),
                            new NamedValue("StyleIdentifier", run.Format.StyleIdentifier),
                            new NamedValue("StyleName", run.Format.StyleName),
                            new NamedValue("Font", run.Format.Font), 
                            new NamedValue("Size", run.Format.Size.ToString(CultureInfo.InvariantCulture))),
                        run.Text.Escape())
                    .AppendLine();
            }
            else
            {
                this.builder.AppendFormat("<Run>{0}</Run>", run.Text.Escape())
                    .AppendLine();
            }
        }

        public override void VisitBookmarkStart(BookmarkStartProxy bookmarkStart)
        {
            this.builder.AppendFormat("<BookmarkStart {0} />", FormatAttributes(new NamedValue("Name", bookmarkStart.Name)))
                .AppendLine();
        }

        public override void VisitBookmarkEnd(BookmarkEndProxy bookmarkEnd)
        {
            this.builder.AppendFormat("<BookmarkEnd {0} />", FormatAttributes(new NamedValue("Name", bookmarkEnd.Name)));
        }

        public override void VisitGroupShapeStart(GroupShapeProxy groupShape)
        {
            this.builder.AppendLine().AppendFormat("<GroupShape {0}>", FormatAttributes(new NamedValue("Name", groupShape.Name)))
                .AppendLine();
        }

        public override void VisitGroupShapeEnd(GroupShapeProxy groupShape)
        {
            this.builder.AppendLine("</GroupShape>");
        }

        public override void VisitContentControlStart(ContentControlProxy contentControl)
        {
            this.builder
                .AppendFormat(
                    "<ContentControl {0}>",
                    FormatAttributes(new NamedValue("Type", contentControl.Type), new NamedValue("Tag", contentControl.Tag)))
                .AppendLine();
        }

        public override void VisitContentControlEnd(ContentControlProxy contentControl)
        {
            this.builder.AppendLine("</ContentControl>");
        }

        public override void VisitShapeStart(ShapeProxy shape)
        {
            this.builder.AppendFormat("<Shape {0}>", FormatAttributes(new NamedValue("Name", shape.Name)));

            if (this.displayOptions.IncludeImages)
            {
                this.builder
                    .AppendLine()
                    .AppendFormat("<Image>{0}</Image>", shape.Image);
            }
            else
            {
                this.builder.AppendLine();
            }
        }

        public override void VisitShapeEnd(ShapeProxy shape)
        {
            this.builder.AppendLine("</Shape>");
        }

        public override void VisitTableStart(TableProxy table)
        {
            this.builder.AppendLine("<Table>");
        }

        public override void VisitTableEnd(TableProxy table)
        {
            this.builder.AppendLine("</Table>");
        }

        public override void VisitRowStart(RowProxy row)
        {
            this.builder.AppendLine("<Row>");
        }

        public override void VisitRowEnd(RowProxy row)
        {
            this.builder.AppendLine("</Row>");
        }

        public override void VisitCellStart(CellProxy cell)
        {
            this.builder.AppendLine("<Cell>");
        }

        public override void VisitCellEnd(CellProxy cell)
        {
            this.builder.AppendLine("</Cell>");
        }

        public override void VisitDrawingMl(DrawingMlProxy drawingMl)
        {
            if (this.displayOptions.IncludeImages)
            {
                this.builder.AppendFormat("<DrawingMl {0}>", FormatAttributes(new NamedValue("Name", drawingMl.Name)))
                    .AppendLine()
                    .AppendFormat("<Image>{0}{1}{0}</Image>", Environment.NewLine, drawingMl.Image)
                    .AppendLine()
                    .AppendLine("</DrawingMl>");
            }
            else
            {
                this.builder
                    .AppendFormat("<DrawingMl {0} />", FormatAttributes(new NamedValue("Name", drawingMl.Name)))
                    .AppendLine();
            }
        }

        public override void VisitFieldStart(FieldStartProxy fieldStart)
        {
            this.builder.AppendLine("<FieldStart />");
        }

        public override void VisitFieldSeparator(FieldSeparatorProxy fieldSeparator)
        {
            this.builder.AppendLine("<FieldSeparator />");
        }

        public override void VisitFieldEnd(FieldEndProxy fieldEnd)
        {
            this.builder.AppendLine("<FieldEnd />");
        }

        private static string FormatAttributes(params NamedValue[] attributeNameValuePairs)
        {
            return string.Join(" ", attributeNameValuePairs.Select(x => x.Name + "=" + x.Value.EncapsulateWithQuotes()));
        }

        private class NamedValue
        {
            public NamedValue(string name, string value)
            {
                this.Name = name;
                this.Value = value;
            }

            public string Name { get; private set; }

            public string Value { get; private set; }
        }
    }
}
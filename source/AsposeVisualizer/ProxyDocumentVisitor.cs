// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProxyDocumentVisitor.cs" company="Philipp Dolder">
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
    using System.Collections.Generic;
    using Aspose.Words;
    using Aspose.Words.Drawing;
    using Aspose.Words.Fields;
    using Aspose.Words.Markup;
    using Aspose.Words.Tables;

    public class ProxyDocumentVisitor : DocumentVisitor
    {
        private static readonly Dictionary<HeaderFooterType, string> HeaderFooterTypes = new Dictionary<HeaderFooterType, string>
        {
            { HeaderFooterType.HeaderPrimary, "Primary" },
            { HeaderFooterType.HeaderFirst, "First" },
            { HeaderFooterType.HeaderEven, "Even" },
            { HeaderFooterType.FooterPrimary, "Primary" },
            { HeaderFooterType.FooterFirst, "First" },
            { HeaderFooterType.FooterEven, "Even" },
        };

        private readonly IProxyFactory proxyFactory;
        private readonly Stack<ICompositeNodeProxy> traversingParents = new Stack<ICompositeNodeProxy>();

        public ProxyDocumentVisitor(IProxyFactory proxyFactory)
        {
            this.proxyFactory = proxyFactory;
        }

        public INodeProxy Root { get; private set; }

        public override VisitorAction VisitDocumentStart(Document document)
        {
            DocumentProxy documentProxy = this.proxyFactory.CreateDocument();
            this.Root = documentProxy;

            this.traversingParents.Push(documentProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDocumentEnd(Document document)
        {
            ICompositeNodeProxy documentProxy = this.traversingParents.Pop();
            
            ProxyDocumentVisitor.EnsureLegalTreeTraversal<DocumentProxy>(documentProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionStart(Section section)
        {
            SectionProxy sectionProxy = this.proxyFactory.CreateSection(section.PageSetup.Orientation.ToString(), section.PageSetup.PaperSize.ToString());

            this.AddToHierarchy(sectionProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionEnd(Section section)
        {
            ICompositeNodeProxy sectionProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<SectionProxy>(sectionProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBodyStart(Body body)
        {
            BodyProxy bodyProxy = this.proxyFactory.CreateBody();

            this.AddToHierarchy(bodyProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBodyEnd(Body body)
        {
            ICompositeNodeProxy bodyProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<BodyProxy>(bodyProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
        {
            if (headerFooter.IsHeader)
            {
                HeaderProxy headerProxy = this.proxyFactory.CreateHeader(HeaderFooterTypes[headerFooter.HeaderFooterType]);

                this.AddToHierarchy(headerProxy);
            }
            else
            {
                FooterProxy footerProxy = this.proxyFactory.CreateFooter(HeaderFooterTypes[headerFooter.HeaderFooterType]);

                this.AddToHierarchy(footerProxy);
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
        {
            ICompositeNodeProxy headerOrFooterProxy = this.traversingParents.Pop();

            if (headerFooter.IsHeader)
            {
                ProxyDocumentVisitor.EnsureLegalTreeTraversal<HeaderProxy>(headerOrFooterProxy);
            }
            else
            {
                ProxyDocumentVisitor.EnsureLegalTreeTraversal<FooterProxy>(headerOrFooterProxy);
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            ParagraphProxy paragraphProxy = this.proxyFactory.CreateParagraph();

            this.AddToHierarchy(paragraphProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            ICompositeNodeProxy paragraphProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<ParagraphProxy>(paragraphProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRun(Run run)
        {
            RunProxy runProxy = this.proxyFactory.CreateRun(run.Text);

            this.AddLeafToHierarchy(runProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
        {
            BookmarkStartProxy bookmarkStartProxy = this.proxyFactory.CreateBookmarkStart(bookmarkStart.Name);

            this.AddLeafToHierarchy(bookmarkStartProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
        {
            BookmarkEndProxy bookmarkEndProxy = this.proxyFactory.CreateBookmarkEnd(bookmarkEnd.Name);

            this.AddLeafToHierarchy(bookmarkEndProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableStart(Table table)
        {
            TableProxy tableProxy = this.proxyFactory.CreateTable();

            this.AddToHierarchy(tableProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableEnd(Table table)
        {
            ICompositeNodeProxy tableProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<TableProxy>(tableProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowStart(Row row)
        {
            RowProxy rowProxy = this.proxyFactory.CreateRow();

            this.AddToHierarchy(rowProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowEnd(Row row)
        {
            ICompositeNodeProxy rowProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<RowProxy>(rowProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellStart(Cell cell)
        {
            CellProxy cellProxy = this.proxyFactory.CreateCell();

            this.AddToHierarchy(cellProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellEnd(Cell cell)
        {
            ICompositeNodeProxy cellProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<CellProxy>(cellProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDrawingML(DrawingML drawingMl)
        {
            DrawingMlProxy drawingMlProxy = this.proxyFactory.CreateDrawingMl();

            this.AddLeafToHierarchy(drawingMlProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldStart(FieldStart fieldStart)
        {
            FieldStartProxy fieldStartProxy = this.proxyFactory.CreateFieldStart();

            this.AddLeafToHierarchy(fieldStartProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
        {
            FieldSeparatorProxy fieldSeparatorProxy = this.proxyFactory.CreateFieldSeparator();

            this.AddLeafToHierarchy(fieldSeparatorProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
        {
            FieldEndProxy fieldEndProxy = this.proxyFactory.CreateFieldEnd();

            this.AddLeafToHierarchy(fieldEndProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitShapeStart(Shape shape)
        {
            ShapeProxy shapeProxy = this.proxyFactory.CreateShape(shape.Name);

            this.AddToHierarchy(shapeProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitShapeEnd(Shape shape)
        {
            ICompositeNodeProxy shapeProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<ShapeProxy>(shapeProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
        {
            GroupShapeProxy groupShapeProxy = this.proxyFactory.CreateGroupShape(groupShape.Name);

            this.AddToHierarchy(groupShapeProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
        {
            ICompositeNodeProxy groupShapeProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<GroupShapeProxy>(groupShapeProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag sdt)
        {
            ContentControlProxy contentControlProxy = this.proxyFactory.CreateContentControl(sdt.SdtType.ToString(), sdt.Tag);

            this.AddToHierarchy(contentControlProxy);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
        {
            ICompositeNodeProxy contentControlProxy = this.traversingParents.Pop();

            ProxyDocumentVisitor.EnsureLegalTreeTraversal<ContentControlProxy>(contentControlProxy);

            return VisitorAction.Continue;
        }

        private static void EnsureLegalTreeTraversal<TProxy>(ICompositeNodeProxy nodeProxy)
        {
            if (!(nodeProxy is TProxy))
            {
                throw new InvalidOperationException("Illegal tree traversal detected: expected " + typeof(TProxy) + " but was " + nodeProxy.GetType());
            }
        }

        private void AddToHierarchy(ICompositeNodeProxy proxy)
        {
            if (this.Root != null)
            {
                this.traversingParents.Peek().Add(proxy);
            }
            else
            {
                this.Root = proxy;
            }

            this.traversingParents.Push(proxy);
        }

        private void AddLeafToHierarchy(INodeProxy leafProxy)
        {
            if (this.Root != null)
            {
                this.traversingParents.Peek().Add(leafProxy);
            }
            else
            {
                this.Root = leafProxy;
            }
        }
    }
}
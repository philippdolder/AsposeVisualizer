// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NodeVisitor.cs" company="Philipp Dolder">
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
    public abstract class NodeVisitor
    {
        public virtual void VisitDocumentStart(DocumentProxy document)
        {
        }

        public virtual void VisitDocumentEnd(DocumentProxy document)
        {
        }

        public virtual void VisitSectionStart(SectionProxy section)
        {
        }

        public virtual void VisitSectionEnd(SectionProxy section)
        {
        }

        public virtual void VisitBodyStart(BodyProxy body)
        {
        }

        public virtual void VisitBodyEnd(BodyProxy body)
        {
        }

        public virtual void VisitParagraphStart(ParagraphProxy paragraph)
        {
        }

        public virtual void VisitParagraphEnd(ParagraphProxy paragraph)
        {
        }

        public virtual void VisitTableStart(TableProxy table)
        {
        }

        public virtual void VisitTableEnd(TableProxy table)
        {
        }

        public virtual void VisitRun(RunProxy run)
        {
        }

        public virtual void VisitBookmarkStart(BookmarkStartProxy bookmarkStart)
        {
        }

        public virtual void VisitBookmarkEnd(BookmarkEndProxy bookmarkEnd)
        {
        }

        public virtual void VisitRowStart(RowProxy row)
        {
        }

        public virtual void VisitRowEnd(RowProxy row)
        {
        }

        public virtual void VisitCellStart(CellProxy cell)
        {
        }

        public virtual void VisitCellEnd(CellProxy cell)
        {
        }

        public virtual void VisitDrawingMl(DrawingMlProxy drawingMl)
        {
        }

        public virtual void VisitFieldStart(FieldStartProxy fieldStart)
        {
        }

        public virtual void VisitFieldSeparator(FieldSeparatorProxy fieldSeparator)
        {
        }

        public virtual void VisitFieldEnd(FieldEndProxy fieldEnd)
        {
        }

        public virtual void VisitShapeStart(ShapeProxy shape)
        {
        }

        public virtual void VisitShapeEnd(ShapeProxy shape)
        {
        }

        public virtual void VisitGroupShapeStart(GroupShapeProxy groupShape)
        {
        }

        public virtual void VisitGroupShapeEnd(GroupShapeProxy groupShape)
        {
        }

        public virtual void VisitContentControlStart(ContentControlProxy contentControl)
        {
        }

        public virtual void VisitContentControlEnd(ContentControlProxy contentControl)
        {
        }

        public virtual void VisitHeaderStart(HeaderProxy header)
        {
        }

        public virtual void VisitHeaderEnd(HeaderProxy header)
        {
        }

        public virtual void VisitFooterStart(FooterProxy footer)
        {
        }

        public virtual void VisitFooterEnd(FooterProxy footer)
        {
        }
    }
}
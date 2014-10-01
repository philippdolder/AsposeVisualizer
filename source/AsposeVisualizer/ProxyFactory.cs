// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProxyFactory.cs" company="Philipp Dolder">
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
    public class ProxyFactory : IProxyFactory
    {
        public DocumentProxy CreateDocument()
        {
            return new DocumentProxy();
        }

        public SectionProxy CreateSection()
        {
            return new SectionProxy();
        }

        public BodyProxy CreateBody()
        {
            return new BodyProxy();
        }

        public ParagraphProxy CreateParagraph()
        {
            return new ParagraphProxy();
        }

        public TableProxy CreateTable()
        {
            return new TableProxy();
        }

        public RunProxy CreateRun(string text)
        {
            return new RunProxy(text);
        }

        public RowProxy CreateRow()
        {
            return new RowProxy();
        }

        public CellProxy CreateCell()
        {
            return new CellProxy();
        }

        public BookmarkStartProxy CreateBookmarkStart(string bookmarkName)
        {
            return new BookmarkStartProxy(bookmarkName);
        }

        public BookmarkEndProxy CreateBookmarkEnd(string bookmarkName)
        {
            return new BookmarkEndProxy(bookmarkName);
        }

        public DrawingMlProxy CreateDrawingMl(string name)
        {
            return new DrawingMlProxy(name);
        }

        public FieldStartProxy CreateFieldStart()
        {
            return new FieldStartProxy();
        }

        public FieldSeparatorProxy CreateFieldSeparator()
        {
            return new FieldSeparatorProxy();
        }

        public FieldEndProxy CreateFieldEnd()
        {
            return new FieldEndProxy();
        }

        public ShapeProxy CreateShape(string shapeName)
        {
            return new ShapeProxy(shapeName);
        }

        public GroupShapeProxy CreateGroupShape(string groupShapeName)
        {
            return new GroupShapeProxy(groupShapeName);
        }

        public ContentControlProxy CreateContentControl(string contentControlType, string tag)
        {
            return new ContentControlProxy(contentControlType, tag);
        }

        public HeaderProxy CreateHeader(string headerType)
        {
            return new HeaderProxy(headerType);
        }

        public FooterProxy CreateFooter(string footerType)
        {
            return new FooterProxy(footerType);
        }
    }
}
// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IProxyFactory.cs" company="Philipp Dolder">
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
    public interface IProxyFactory
    {
        DocumentProxy CreateDocument();

        SectionProxy CreateSection();

        BodyProxy CreateBody();

        ParagraphProxy CreateParagraph();

        TableProxy CreateTable();

        RunProxy CreateRun(string text);

        RowProxy CreateRow();

        CellProxy CreateCell();

        BookmarkStartProxy CreateBookmarkStart(string bookmarkName);

        BookmarkEndProxy CreateBookmarkEnd(string bookmarkName);

        DrawingMlProxy CreateDrawingMl();

        FieldStartProxy CreateFieldStart();

        FieldSeparatorProxy CreateFieldSeparator();

        FieldEndProxy CreateFieldEnd();

        ShapeProxy CreateShape(string shapeName);

        GroupShapeProxy CreateGroupShape(string groupShapeName);

        ContentControlProxy CreateContentControl(string contentControlType, string tag);

        HeaderProxy CreateHeader(string headerType);

        FooterProxy CreateFooter(string footerType);
    }
}
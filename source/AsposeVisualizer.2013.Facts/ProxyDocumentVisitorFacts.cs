// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProxyDocumentVisitorFacts.cs" company="Philipp Dolder">
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
    using Aspose.Words;
    using Aspose.Words.Tables;
    using FakeItEasy;
    using FluentAssertions;
    using Xunit;

    public class ProxyDocumentVisitorFacts
    {
        private readonly ProxyDocumentVisitor testee;
        private readonly IProxyFactory proxyFactory;

        public ProxyDocumentVisitorFacts()
        {
            this.proxyFactory = A.Fake<IProxyFactory>();

            this.testee = new ProxyDocumentVisitor(this.proxyFactory);
        }

        [Fact]
        public void SetsDocumentAsRoot_WhenStartVisitingFromDocument()
        {
            var document = new Document();
            var documentProxy = new DocumentProxy();

            A.CallTo(() => this.proxyFactory.CreateDocument()).Returns(documentProxy);

            document.Accept(this.testee);

            this.testee.Root.Should().Be(documentProxy);
        }

        [Fact]
        public void AddsSectionsToDocument_WhenStartVisitingFromDocument()
        {
            Document document = CreateDocumentWithTwoSections();

            var documentProxy = new DocumentProxy();
            var firstSectionProxy = A.Fake<SectionProxy>();
            var secondSectionProxy = A.Fake<SectionProxy>();

            A.CallTo(() => this.proxyFactory.CreateDocument()).Returns(documentProxy);
            A.CallTo(() => this.proxyFactory.CreateSection("Portrait", "Letter")).Returns(firstSectionProxy);
            A.CallTo(() => this.proxyFactory.CreateSection("Landscape", "Letter")).Returns(secondSectionProxy);

            document.Accept(this.testee);

            documentProxy.Sections
                .Should().HaveCount(2)
                .And.ContainInOrder(firstSectionProxy, secondSectionProxy);
        }

        [Fact]
        public void AddsBodyToSections_WhenStartVisitingFromDocument()
        {
            Document document = CreateDocumentWithTwoSections();

            var firstSectionProxy = new SectionProxy("Orientation", "PaperSize");
            var secondSectionProxy = new SectionProxy("Orientation", "PaperSize");
            var firstBodyProxy = A.Fake<BodyProxy>();
            var secondBodyProxy = A.Fake<BodyProxy>();

            A.CallTo(() => this.proxyFactory.CreateSection(A<string>._, A<string>._)).ReturnsNextFromSequence(firstSectionProxy, secondSectionProxy);
            A.CallTo(() => this.proxyFactory.CreateBody()).ReturnsNextFromSequence(firstBodyProxy, secondBodyProxy);

            document.Accept(this.testee);

            firstSectionProxy.Body.Should().Be(firstBodyProxy);
            secondSectionProxy.Body.Should().Be(secondBodyProxy);
        }

        [Fact]
        public void SetsSectionAsRoot_WhenStartVisitingFromSection()
        {
            Section section = new Document().FirstSection;
            var sectionProxy = A.Fake<SectionProxy>();

            A.CallTo(() => this.proxyFactory.CreateSection(A<string>._, A<string>._)).Returns(sectionProxy);

            section.Accept(this.testee);

            this.testee.Root.Should().Be(sectionProxy);
        }

        [Fact]
        public void AddsBodyToSection_WhenStartVisitingFromSection()
        {
            Section section = new Document().FirstSection;
            var sectionProxy = new SectionProxy("Orientation", "PaperSize");
            var bodyProxy = A.Fake<BodyProxy>();

            A.CallTo(() => this.proxyFactory.CreateSection(A<string>._, A<string>._)).Returns(sectionProxy);
            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);

            section.Accept(this.testee);

            sectionProxy.Body.Should().Be(bodyProxy);
        }

        [Fact]
        public void AddsHeadersToSection_WhenStartVisitingFromSection()
        {
            var builder = new DocumentBuilder();
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);

            Section section = builder.Document.FirstSection;

            var sectionProxy = new SectionProxy("Orientation", "PaperSize");
            var firstHeaderProxy = A.Fake<HeaderProxy>();
            var secondHeaderProxy = A.Fake<HeaderProxy>();

            A.CallTo(() => this.proxyFactory.CreateSection(A<string>._, A<string>._)).Returns(sectionProxy);
            A.CallTo(() => this.proxyFactory.CreateHeader("Primary")).Returns(firstHeaderProxy);
            A.CallTo(() => this.proxyFactory.CreateHeader("Even")).Returns(secondHeaderProxy);

            section.Accept(this.testee);

            sectionProxy.Headers.Should().Contain(firstHeaderProxy)
                .And.Contain(secondHeaderProxy);
        }

        [Fact]
        public void AddsFootersToSection_WhenStartVisitingFromSection()
        {
            var builder = new DocumentBuilder();
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            Section section = builder.Document.FirstSection;

            var sectionProxy = new SectionProxy("Orientation", "PaperSize");
            var firstFooterProxy = A.Fake<FooterProxy>();
            var secondFooterProxy = A.Fake<FooterProxy>();

            A.CallTo(() => this.proxyFactory.CreateSection(A<string>._, A<string>._)).Returns(sectionProxy);
            A.CallTo(() => this.proxyFactory.CreateFooter("Primary")).Returns(firstFooterProxy);
            A.CallTo(() => this.proxyFactory.CreateFooter("Even")).Returns(secondFooterProxy);

            section.Accept(this.testee);

            sectionProxy.Footers.Should().Contain(firstFooterProxy)
                .And.Contain(secondFooterProxy);
        }

        [Fact]
        public void AddsParagraphsToBody_WhenStartVisitingFromSection()
        {
            var builder = new DocumentBuilder();
            builder.InsertParagraph();
            Section section = builder.Document.FirstSection;
            var bodyProxy = new BodyProxy();
            var firstParagraph = A.Fake<ParagraphProxy>();
            var secondParagraph = A.Fake<ParagraphProxy>();

            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);
            A.CallTo(() => this.proxyFactory.CreateParagraph()).ReturnsNextFromSequence(firstParagraph, secondParagraph);

            section.Accept(this.testee);

            bodyProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstParagraph, secondParagraph);
        }

        [Fact]
        public void AddsTablesToBody_WhenStartVisitingFromSection()
        {
            Section section = CreateDocumentWithTables(2, 1, 1).FirstSection;
            var bodyProxy = new BodyProxy();
            var firstTableProxy = A.Fake<TableProxy>();
            var secondTableProxy = A.Fake<TableProxy>();

            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);
            A.CallTo(() => this.proxyFactory.CreateTable()).ReturnsNextFromSequence(firstTableProxy, secondTableProxy);

            section.Accept(this.testee);

            bodyProxy.Children.Should().ContainInOrder(firstTableProxy, secondTableProxy);
        }

        [Fact]
        public void SetsBodyAsRoot_WhenStartVisitingFromBody()
        {
            Body body = new Document().FirstSection.Body;
            var bodyProxy = A.Fake<BodyProxy>();

            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);

            body.Accept(this.testee);

            this.testee.Root.Should().Be(bodyProxy);
        }

        [Fact]
        public void AddsParagraphsToBody_WhenStartVisitingFromBody()
        {
            var builder = new DocumentBuilder();
            builder.InsertParagraph();
            Body body = builder.Document.FirstSection.Body;

            var bodyProxy = new BodyProxy();
            var firstParagraphProxy = A.Fake<ParagraphProxy>();
            var secondParagraphProxy = A.Fake<ParagraphProxy>();

            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);
            A.CallTo(() => this.proxyFactory.CreateParagraph()).ReturnsNextFromSequence(firstParagraphProxy, secondParagraphProxy);

            body.Accept(this.testee);

            bodyProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstParagraphProxy, secondParagraphProxy);
        }

        [Fact]
        public void AddsRunsToParagraph_WhenStartVisitingFromBody()
        {
            var builder = new DocumentBuilder();
            builder.Write("FirstRun");
            builder.Write("SecondRun");
            Body body = builder.Document.FirstSection.Body;

            var paragraphProxy = new ParagraphProxy();
            var firstRunProxy = A.Fake<RunProxy>();
            var secondRunProxy = A.Fake<RunProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateRun("FirstRun")).Returns(firstRunProxy);
            A.CallTo(() => this.proxyFactory.CreateRun("SecondRun")).Returns(secondRunProxy);

            body.Accept(this.testee);

            paragraphProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstRunProxy, secondRunProxy);
        }

        [Fact]
        public void AddsTablesToBody_WhenStartVisitingFromBody()
        {
            Body body = CreateDocumentWithTables(2, 1, 1).FirstSection.Body;

            var bodyProxy = new BodyProxy();
            var firstTableProxy = A.Fake<TableProxy>();
            var secondTableProxy = A.Fake<TableProxy>();

            A.CallTo(() => this.proxyFactory.CreateBody()).Returns(bodyProxy);
            A.CallTo(() => this.proxyFactory.CreateTable()).ReturnsNextFromSequence(firstTableProxy, secondTableProxy);

            body.Accept(this.testee);

            bodyProxy.Children.Should().ContainInOrder(firstTableProxy, secondTableProxy);
        }

        [Fact]
        public void AddsRowsToTable_WhenStartVisitingFromBody()
        {
            Body body = CreateDocumentWithTables(1, 2, 1).FirstSection.Body;

            var tableProxy = new TableProxy();
            var firstRowProxy = A.Fake<RowProxy>();
            var secondRowProxy = A.Fake<RowProxy>();

            A.CallTo(() => this.proxyFactory.CreateTable()).Returns(tableProxy);
            A.CallTo(() => this.proxyFactory.CreateRow()).ReturnsNextFromSequence(firstRowProxy, secondRowProxy);

            body.Accept(this.testee);

            tableProxy.Rows.Should().HaveCount(2)
                .And.ContainInOrder(firstRowProxy, secondRowProxy);
        }

        [Fact]
        public void SetsParagraphAsRoot_WhenStartVisitingFromParagraph()
        {
            Paragraph paragraph = new Document().FirstSection.Body.FirstParagraph;

            var paragraphProxy = A.Fake<ParagraphProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);

            paragraph.Accept(this.testee);

            this.testee.Root.Should().Be(paragraphProxy);
        }

        [Fact]
        public void AddsRunsToParagraph_WhenStartVisitingFromParagraph()
        {
            var builder = new DocumentBuilder();
            builder.Write("FirstRun");
            builder.Write("SecondRun");
            Paragraph paragraph = builder.Document.FirstSection.Body.FirstParagraph;

            var paragraphProxy = new ParagraphProxy();
            var firstRunProxy = A.Fake<RunProxy>();
            var secondRunProxy = A.Fake<RunProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateRun(A<string>._)).ReturnsNextFromSequence(firstRunProxy, secondRunProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstRunProxy, secondRunProxy);
        }

        [Fact]
        public void AddsBookmarkStartsToParagraph_WhenStartVisitingFromParagraph()
        {
            const string FirstBookmarkName = "Bookmark 1";
            const string SecondBookmarkName = "Bookmark 2";
            var builder = new DocumentBuilder();
            builder.StartBookmark(FirstBookmarkName);
            builder.EndBookmark(FirstBookmarkName);
            builder.StartBookmark(SecondBookmarkName);
            builder.EndBookmark(SecondBookmarkName);
            Paragraph paragraph = builder.Document.FirstSection.Body.FirstParagraph;

            var paragraphProxy = new ParagraphProxy();
            var firstBookmarkStartProxy = A.Fake<BookmarkStartProxy>();
            var secondBookmarkStartProxy = A.Fake<BookmarkStartProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateBookmarkStart(FirstBookmarkName)).Returns(firstBookmarkStartProxy);
            A.CallTo(() => this.proxyFactory.CreateBookmarkStart(SecondBookmarkName)).Returns(secondBookmarkStartProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().ContainInOrder(firstBookmarkStartProxy, secondBookmarkStartProxy);
        }

        [Fact]
        public void AddsBookmarkEndsToParagraph_WhenStartVisitingFromParagraph()
        {
            const string FirstBookmarkName = "Bookmark 1";
            const string SecondBookmarkName = "Bookmark 2";
            var builder = new DocumentBuilder();
            builder.StartBookmark(FirstBookmarkName);
            builder.EndBookmark(FirstBookmarkName);
            builder.StartBookmark(SecondBookmarkName);
            builder.EndBookmark(SecondBookmarkName);
            Paragraph paragraph = builder.Document.FirstSection.Body.FirstParagraph;

            var paragraphProxy = new ParagraphProxy();
            var firstBookmarkEndProxy = A.Fake<BookmarkEndProxy>();
            var secondBookmarkEndProxy = A.Fake<BookmarkEndProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateBookmarkEnd(FirstBookmarkName)).Returns(firstBookmarkEndProxy);
            A.CallTo(() => this.proxyFactory.CreateBookmarkEnd(SecondBookmarkName)).Returns(secondBookmarkEndProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().ContainInOrder(firstBookmarkEndProxy, secondBookmarkEndProxy);
        }

        [Fact]
        public void AddsFieldStartToParagraph_WhenStartVisitingFromParagraph()
        {
            var builder = new DocumentBuilder();
            builder.InsertField("PAGE");

            Paragraph paragraph = builder.Document.FirstSection.Body.LastParagraph;

            var paragraphProxy = new ParagraphProxy();
            var fieldStartProxy = A.Fake<FieldStartProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateFieldStart()).Returns(fieldStartProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().Contain(fieldStartProxy);
        }

        [Fact]
        public void AddsFieldEndToParagraph_WhenStartVisitingFromParagraph()
        {
            var builder = new DocumentBuilder();
            builder.InsertField("PAGE");

            Paragraph paragraph = builder.Document.FirstSection.Body.LastParagraph;

            var paragraphProxy = new ParagraphProxy();
            var fieldEndProxy = A.Fake<FieldEndProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateFieldEnd()).Returns(fieldEndProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().Contain(fieldEndProxy);
        }

        [Fact]
        public void AddsFieldSeparatorToParagraph_WhenStartVisitingFromParagraph()
        {
            var builder = new DocumentBuilder();
            builder.InsertField("PAGE");

            Paragraph paragraph = builder.Document.FirstSection.Body.LastParagraph;

            var paragraphProxy = new ParagraphProxy();
            var fieldSeparatorProxy = A.Fake<FieldSeparatorProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateFieldSeparator()).Returns(fieldSeparatorProxy);

            paragraph.Accept(this.testee);

            paragraphProxy.Children.Should().Contain(fieldSeparatorProxy);
        }

        [Fact]
        public void SetsBookmarkStartAsRoot_WhenStartVisitingFromBookmarkStart()
        {
            const string BookmarkName = "Bookmark";
            var builder = new DocumentBuilder();
            BookmarkStart bookmarkStart = builder.StartBookmark(BookmarkName);
            builder.EndBookmark(BookmarkName);

            var bookmarkStartProxy = A.Fake<BookmarkStartProxy>();

            A.CallTo(() => this.proxyFactory.CreateBookmarkStart(BookmarkName)).Returns(bookmarkStartProxy);

            bookmarkStart.Accept(this.testee);

            this.testee.Root.Should().Be(bookmarkStartProxy);
        }

        [Fact]
        public void SetsBookmarkEndAsRoot_WhenStartVisitingFromBookmarkEnd()
        {
            const string BookmarkName = "Bookmark";
            var builder = new DocumentBuilder();
            builder.StartBookmark(BookmarkName);
            BookmarkEnd bookmarkEnd = builder.EndBookmark(BookmarkName);

            var bookmarkEndProxy = A.Fake<BookmarkEndProxy>();

            A.CallTo(() => this.proxyFactory.CreateBookmarkEnd(BookmarkName)).Returns(bookmarkEndProxy);

            bookmarkEnd.Accept(this.testee);

            this.testee.Root.Should().Be(bookmarkEndProxy);
        }

        [Fact]
        public void SetsTableAsRoot_WhenStartVisitingFromTable()
        {
            Table table = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0];

            var tableProxy = new TableProxy();

            A.CallTo(() => this.proxyFactory.CreateTable()).Returns(tableProxy);

            table.Accept(this.testee);

            this.testee.Root.Should().Be(tableProxy);
        }

        [Fact]
        public void AddsRowsToTable_WhenStartVisitingFromTable()
        {
            Table table = CreateDocumentWithTables(1, 2, 1).FirstSection.Body.Tables[0];

            var tableProxy = new TableProxy();
            var firstRowProxy = A.Fake<RowProxy>();
            var secondRowProxy = A.Fake<RowProxy>();

            A.CallTo(() => this.proxyFactory.CreateTable()).Returns(tableProxy);
            A.CallTo(() => this.proxyFactory.CreateRow()).ReturnsNextFromSequence(firstRowProxy, secondRowProxy);

            table.Accept(this.testee);

            tableProxy.Rows.Should().HaveCount(2)
                .And.ContainInOrder(firstRowProxy, secondRowProxy);
        }

        [Fact]
        public void AddsCellsToRow_WhenStartVisitingFromTable()
        {
            Table table = CreateDocumentWithTables(1, 1, 2).FirstSection.Body.Tables[0];

            var rowProxy = new RowProxy();
            var firstCellProxy = A.Fake<CellProxy>();
            var secondCellProxy = A.Fake<CellProxy>();

            A.CallTo(() => this.proxyFactory.CreateRow()).Returns(rowProxy);
            A.CallTo(() => this.proxyFactory.CreateCell()).ReturnsNextFromSequence(firstCellProxy, secondCellProxy);

            table.Accept(this.testee);

            rowProxy.Cells.Should().HaveCount(2)
                .And.ContainInOrder(firstCellProxy, secondCellProxy);
        }

        [Fact]
        public void AddsInnerTablesToCell_WhenStartVisitingFromOuterTable()
        {
            Table outerTable = CreateDocumentWithNestedTables();

            var outerCellProxy = new CellProxy();
            var firstInnerTableProxy = A.Fake<TableProxy>();
            var secondInnerTableProxy = A.Fake<TableProxy>();

            A.CallTo(() => this.proxyFactory.CreateCell()).ReturnsNextFromSequence(outerCellProxy);
            A.CallTo(() => this.proxyFactory.CreateTable()).ReturnsNextFromSequence(A.Fake<TableProxy>(), firstInnerTableProxy, secondInnerTableProxy);

            outerTable.Accept(this.testee);

            outerCellProxy.Children.Should().ContainInOrder(firstInnerTableProxy, secondInnerTableProxy);
        }

        [Fact]
        public void SetsRowAsRoot_WhenStartVisitingFromRow()
        {
            Row row = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0].FirstRow;

            var rowProxy = A.Fake<RowProxy>();

            A.CallTo(() => this.proxyFactory.CreateRow()).Returns(rowProxy);

            row.Accept(this.testee);

            this.testee.Root.Should().Be(rowProxy);
        }

        [Fact]
        public void AddsCellsToRow_WhenStartVisitingFromRow()
        {
            Row row = CreateDocumentWithTables(1, 1, 2).FirstSection.Body.Tables[0].FirstRow;

            var rowProxy = new RowProxy();
            var firstCellProxy = A.Fake<CellProxy>();
            var secondCellProxy = A.Fake<CellProxy>();

            A.CallTo(() => this.proxyFactory.CreateRow()).Returns(rowProxy);
            A.CallTo(() => this.proxyFactory.CreateCell()).ReturnsNextFromSequence(firstCellProxy, secondCellProxy);

            row.Accept(this.testee);

            rowProxy.Cells.Should().HaveCount(2)
                .And.ContainInOrder(firstCellProxy, secondCellProxy);
        }

        [Fact]
        public void AddsParagraphsToCell_WhenStartVisitingFromRow()
        {
            Row row = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0].FirstRow;
            var secondParagraph = new Paragraph(row.Document);
            row.FirstCell.AppendChild(secondParagraph);

            var cellProxy = new CellProxy();
            var firstParagraphProxy = new ParagraphProxy();
            var secondParagraphProxy = new ParagraphProxy();

            A.CallTo(() => this.proxyFactory.CreateCell()).Returns(cellProxy);
            A.CallTo(() => this.proxyFactory.CreateParagraph()).ReturnsNextFromSequence(firstParagraphProxy, secondParagraphProxy);

            row.Accept(this.testee);

            cellProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstParagraphProxy, secondParagraphProxy);
        }

        [Fact]
        public void SetsCellAsRoot_WhenStartVisitingFromCell()
        {
            Cell cell = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0].FirstRow.FirstCell;
            var cellProxy = A.Fake<CellProxy>();

            A.CallTo(() => this.proxyFactory.CreateCell()).Returns(cellProxy);

            cell.Accept(this.testee);

            this.testee.Root.Should().Be(cellProxy);
        }

        [Fact]
        public void AddsParagraphsToCell_WhenStartVisitingFromCell()
        {
            Cell cell = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0].FirstRow.FirstCell;
            var secondParagraph = new Paragraph(cell.Document);
            cell.AppendChild(secondParagraph);

            var cellProxy = new CellProxy();
            var firstParagraphProxy = new ParagraphProxy();
            var secondParagraphProxy = new ParagraphProxy();

            A.CallTo(() => this.proxyFactory.CreateCell()).Returns(cellProxy);
            A.CallTo(() => this.proxyFactory.CreateParagraph()).ReturnsNextFromSequence(firstParagraphProxy, secondParagraphProxy);

            cell.Accept(this.testee);

            cellProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstParagraphProxy, secondParagraphProxy);
        }

        [Fact]
        public void AddsRunsToParagraph_WhenStartVisitingFromCell()
        {
            Cell cell = CreateDocumentWithTables(1, 1, 1).FirstSection.Body.Tables[0].FirstRow.FirstCell;
            var firstRun = new Run(cell.Document);
            var secondRun = new Run(cell.Document);
            cell.FirstParagraph.AppendChild(firstRun);
            cell.FirstParagraph.AppendChild(secondRun);

            var paragraphProxy = new ParagraphProxy();
            var firstRunProxy = A.Fake<RunProxy>();
            var secondRunProxy = A.Fake<RunProxy>();

            A.CallTo(() => this.proxyFactory.CreateParagraph()).Returns(paragraphProxy);
            A.CallTo(() => this.proxyFactory.CreateRun(A<string>._)).ReturnsNextFromSequence(firstRunProxy, secondRunProxy);

            cell.Accept(this.testee);

            paragraphProxy.Children.Should().HaveCount(2)
                .And.ContainInOrder(firstRunProxy, secondRunProxy);
        }

        [Fact]
        public void SetsRunAsRoot_WhenStartVisitingFromRun()
        {
            var builder = new DocumentBuilder();
            const string Text = "Text";
            builder.Write(Text);
            Run run = builder.Document.FirstSection.Body.FirstParagraph.Runs[0];

            var runProxy = A.Fake<RunProxy>();

            A.CallTo(() => this.proxyFactory.CreateRun(Text)).Returns(runProxy);

            run.Accept(this.testee);

            this.testee.Root.Should().Be(runProxy);
        }

        private static Document CreateDocumentWithTwoSections()
        {
            var builder = new DocumentBuilder();
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.CurrentSection.PageSetup.Orientation = Orientation.Landscape;

            return builder.Document;
        }

        private static Document CreateDocumentWithTables(int numberOfTables, int numberOfRows, int numberOfCells)
        {
            var builder = new DocumentBuilder();

            for (int i = 0; i < numberOfTables; i++)
            {
                builder.StartTable();

                for (int j = 0; j < numberOfRows; j++)
                {
                    for (int k = 0; k < numberOfCells; k++)
                    {
                        builder.InsertCell();
                    }

                    builder.EndRow();
                }

                builder.EndTable();
            }

            return builder.Document;
        }

        private static Table CreateDocumentWithNestedTables()
        {
            var builder = new DocumentBuilder();
            Table outerTable = builder.StartTable();
            var outerCell = builder.InsertCell();
            builder.EndRow();
            builder.EndTable();

            builder.MoveTo(outerCell.FirstParagraph);
            builder.StartTable();
            builder.InsertCell();
            builder.EndRow();
            builder.EndTable();

            builder.StartTable();
            builder.InsertCell();
            builder.EndRow();
            builder.EndTable();

            return outerTable;
        }
    }
}
// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ParagraphProxyFacts.cs" company="Philipp Dolder">
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
    using FakeItEasy;
    using FluentAssertions;
    using Xunit;

    public class ParagraphProxyFacts
    {
        private readonly ParagraphProxy testee;

        public ParagraphProxyFacts()
        {
            this.testee = new ParagraphProxy();
        }

        [Fact]
        public void HasChildrenInAddedOrder()
        {
            var firstRun = new RunProxy("text");
            var secondRun = new RunProxy("text");

            this.testee.Add(firstRun);
            this.testee.Add(secondRun);

            this.testee.Children.Should().ContainInOrder(firstRun, secondRun);
        }

        [Fact]
        public void ChildCanBeRun()
        {
            var run = new RunProxy("text");

            this.testee.Add(run);

            this.testee.Children.Should().Contain(run);
        }

        [Fact]
        public void ChildCanBeBookmarkStart()
        {
            var bookmarkStart = new BookmarkStartProxy("name");

            this.testee.Add(bookmarkStart);

            this.testee.Children.Should().Contain(bookmarkStart);
        }

        [Fact]
        public void ChildCanBeBookmarkEnd()
        {
            var bookmarkEnd = new BookmarkEndProxy("name");

            this.testee.Add(bookmarkEnd);

            this.testee.Children.Should().Contain(bookmarkEnd);
        }

        [Fact]
        public void ChildCanBeDrawingMl()
        {
            var drawingMl = new DrawingMlProxy("name");

            this.testee.Add(drawingMl);

            this.testee.Children.Should().Contain(drawingMl);
        }

        [Fact]
        public void HasFormatting()
        {
            this.testee.Format.Should().NotBeNull();
        }

        [Fact]
        public void VisitsChildrenBetweenParagraphStartAndEnd()
        {
            var child = A.Fake<INodeProxy>();
            this.testee.Add(child);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitParagraphStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => child.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitParagraphEnd(this.testee)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void SendsVisitorToChildrenInOrder()
        {
            var bookmarkStart = A.Fake<BookmarkStartProxy>();
            var text = A.Fake<RunProxy>();
            var bookmarkEnd = A.Fake<BookmarkEndProxy>();

            this.testee.Add(bookmarkStart);
            this.testee.Add(text);
            this.testee.Add(bookmarkEnd);

            var visitor = A.Fake<NodeVisitor>();

            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => bookmarkStart.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => text.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => bookmarkEnd.Accept(visitor)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void IsSerializable()
        {
            this.testee.GetType().Should().BeDecoratedWith<SerializableAttribute>();
        }
    }
}
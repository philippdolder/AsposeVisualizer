﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ContentControlProxyFacts.cs" company="Philipp Dolder">
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

    public class ContentControlProxyFacts
    {
        private const string ContentControlType = "Type";
        private const string Tag = "Tag";

        private readonly ContentControlProxy testee;

        public ContentControlProxyFacts()
        {
            this.testee = new ContentControlProxy(ContentControlType, Tag);
        }

        [Fact]
        public void HasTag()
        {
            this.testee.Tag.Should().Be(Tag);
        }

        [Fact]
        public void HasContentControlType()
        {
            this.testee.Type.Should().Be(ContentControlType);
        }

        [Fact]
        public void HasChildrenInAddedOrder()
        {
            var firstChild = A.Fake<INodeProxy>();
            var secondChild = A.Fake<INodeProxy>();

            this.testee.Add(firstChild);
            this.testee.Add(secondChild);

            this.testee.Children.Should().ContainInOrder(firstChild, secondChild);
        }

        [Fact]
        public void VisitsChildrenBetweenContentControlStartAndEnd()
        {
            var child = A.Fake<INodeProxy>();

            this.testee.Add(child);

            var visitor = A.Fake<NodeVisitor>();

            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitContentControlStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => child.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitContentControlEnd(this.testee)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void SendsVisitorToChildrenInOrder()
        {
            var firstChild = A.Fake<INodeProxy>();
            var secondChild = A.Fake<INodeProxy>();

            this.testee.Add(firstChild);
            this.testee.Add(secondChild);

            var visitor = A.Fake<NodeVisitor>();

            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => firstChild.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => secondChild.Accept(visitor)).MustHaveHappened();
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
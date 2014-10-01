// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TableProxyFacts.cs" company="Philipp Dolder">
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

    public class TableProxyFacts
    {
        private readonly TableProxy testee;

        public TableProxyFacts()
        {
            this.testee = new TableProxy();
        }

        [Fact]
        public void HasRowsInAddedOrder()
        {
            var firstRow = new RowProxy();
            var secondRow = new RowProxy();

            this.testee.Add(firstRow);
            this.testee.Add(secondRow);

            this.testee.Rows.Should().ContainInOrder(firstRow, secondRow);
        }

        [Fact]
        public void SendsVisitorToRowsInOrder()
        {
            var firstRow = A.Fake<RowProxy>();
            var secondRow = A.Fake<RowProxy>();
            this.testee.Add(firstRow);
            this.testee.Add(secondRow);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => firstRow.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => secondRow.Accept(visitor)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void VisitsRowsBetweenTableStartAndEnd()
        {
            var row = A.Fake<RowProxy>();
            this.testee.Add(row);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitTableStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => row.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitTableEnd(this.testee)).MustHaveHappened();
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
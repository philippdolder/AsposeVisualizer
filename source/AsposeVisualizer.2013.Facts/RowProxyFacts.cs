// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RowProxyFacts.cs" company="Philipp Dolder">
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

    public class RowProxyFacts
    {
        private readonly RowProxy testee;

        public RowProxyFacts()
        {
            this.testee = new RowProxy();
        }

        [Fact]
        public void HasCellsInAddedOrder()
        {
            var firstCell = A.Fake<CellProxy>();
            var secondCell = A.Fake<CellProxy>();

            this.testee.Add(firstCell);
            this.testee.Add(secondCell);

            this.testee.Cells.Should().ContainInOrder(firstCell, secondCell);
        }

        [Fact]
        public void SendsVisitorToCellsInOrder()
        {
            var firstCell = A.Fake<CellProxy>();
            var secondCell = A.Fake<CellProxy>();
            this.testee.Add(firstCell);
            this.testee.Add(secondCell);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => firstCell.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => secondCell.Accept(visitor)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void VisitsCellsBetweenRowStartAndEnd()
        {
            var cell = A.Fake<CellProxy>();
            this.testee.Add(cell);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitRowStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => cell.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitRowEnd(this.testee)).MustHaveHappened();
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
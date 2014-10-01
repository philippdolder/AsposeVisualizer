// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocumentProxyFacts.cs" company="Philipp Dolder">
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

    public class DocumentProxyFacts
    {
        private readonly DocumentProxy testee;

        public DocumentProxyFacts()
        {
            this.testee = new DocumentProxy();
        }

        [Fact]
        public void SendsVisitorToSectionsInOrder()
        {
            var firstSection = A.Fake<SectionProxy>();
            var secondSection = A.Fake<SectionProxy>();
            this.testee.Add(firstSection);
            this.testee.Add(secondSection);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => firstSection.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => secondSection.Accept(visitor)).MustHaveHappened();
                }
            }
        }

        [Fact]
        public void VisitsSectionsBetweenDocumentStartAndEnd()
        {
            var section = A.Fake<SectionProxy>();
            this.testee.Add(section);

            var visitor = A.Fake<NodeVisitor>();
            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitDocumentStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => section.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitDocumentEnd(this.testee)).MustHaveHappened();
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
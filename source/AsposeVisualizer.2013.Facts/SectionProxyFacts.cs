// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SectionProxyFacts.cs" company="Philipp Dolder">
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

    public class SectionProxyFacts
    {
        private const string Orientation = "orientation";
        private const string PaperSize = "paperSize";
        private readonly SectionProxy testee;

        public SectionProxyFacts()
        {
            this.testee = new SectionProxy(Orientation, PaperSize);
        }

        [Fact]
        public void HasOrientation()
        {
            this.testee.Orientation.Should().Be(Orientation);
        }

        [Fact]
        public void HasPaperSize()
        {
            this.testee.PaperSize.Should().Be(PaperSize);
        }

        [Fact]
        public void HasBody()
        {
            var body = A.Fake<BodyProxy>();
            this.testee.Add(body);

            this.testee.Body.Should().Be(body);
        }

        [Fact]
        public void SendsVisitorToBody()
        {
            var body = A.Fake<BodyProxy>();
            this.testee.Add(body);

            var visitor = A.Fake<NodeVisitor>();
            this.testee.Accept(visitor);

            A.CallTo(() => body.Accept(visitor)).MustHaveHappened();
        }

        [Fact]
        public void VisitsBodyBetweenSectionStartAndEnd()
        {
            var body = A.Fake<BodyProxy>();

            this.testee.Add(body);

            var visitor = A.Fake<NodeVisitor>();

            using (var scope = Fake.CreateScope())
            {
                this.testee.Accept(visitor);

                using (scope.OrderedAssertions())
                {
                    A.CallTo(() => visitor.VisitSectionStart(this.testee)).MustHaveHappened();
                    A.CallTo(() => body.Accept(visitor)).MustHaveHappened();
                    A.CallTo(() => visitor.VisitSectionEnd(this.testee)).MustHaveHappened();
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
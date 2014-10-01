// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocumentProxy.cs" company="Philipp Dolder">
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
    using System.Collections.Generic;

    [Serializable]
    public class DocumentProxy : ICompositeNodeProxy
    {
        private readonly List<SectionProxy> sections = new List<SectionProxy>();

        public IReadOnlyList<SectionProxy> Sections
        {
            get { return this.sections; }
        }

        public void Add(INodeProxy node)
        {
            this.sections.Add((SectionProxy)node);
        }

        public void Accept(NodeVisitor visitor)
        {
            visitor.VisitDocumentStart(this);

            foreach (SectionProxy section in this.sections)
            {
                section.Accept(visitor);
            }

            visitor.VisitDocumentEnd(this);
        }
    }
}
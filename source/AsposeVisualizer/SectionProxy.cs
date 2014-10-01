// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SectionProxy.cs" company="Philipp Dolder">
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
    public class SectionProxy : ICompositeNodeProxy
    {
        private readonly List<HeaderProxy> headers = new List<HeaderProxy>();
        private readonly List<FooterProxy> footers = new List<FooterProxy>();

        public SectionProxy()
        {
            this.Format = new SectionFormatProxy();
        }

        public BodyProxy Body { get; private set; }

        public IReadOnlyCollection<HeaderProxy> Headers
        {
            get { return this.headers; }
        }

        public IReadOnlyCollection<FooterProxy> Footers
        {
            get { return this.footers; }
        }

        public SectionFormatProxy Format { get; private set; }

        public void Add(INodeProxy node)
        {
            if (node is BodyProxy)
            {
                this.Body = (BodyProxy)node;
            }
            else if (node is HeaderProxy)
            {
                this.headers.Add((HeaderProxy)node);
            }
            else if (node is FooterProxy)
            {
                this.footers.Add((FooterProxy)node);
            }
            else
            {
                throw new InvalidOperationException("node of type " + node.GetType() + " is not allowed as a child of Section");
            }
        }

        public virtual void Accept(NodeVisitor visitor)
        {
            visitor.VisitSectionStart(this);

            this.Body.Accept(visitor);

            visitor.VisitSectionEnd(this);
        }
    }
}
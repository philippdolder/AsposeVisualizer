// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellProxy.cs" company="Philipp Dolder">
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
    public class CellProxy : ICompositeNodeProxy
    {
        private readonly List<INodeProxy> children = new List<INodeProxy>();

        public IReadOnlyList<INodeProxy> Children
        {
            get { return this.children; }
        }

        public virtual void Accept(NodeVisitor visitor)
        {
            visitor.VisitCellStart(this);

            foreach (INodeProxy child in this.children)
            {
                child.Accept(visitor);
            }

            visitor.VisitCellEnd(this);
        }

        public void Add(INodeProxy node)
        {
            this.children.Add(node);
        }
    }
}
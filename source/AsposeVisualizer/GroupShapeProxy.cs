// --------------------------------------------------------------------------------------------------------------------
// <copyright file="GroupShapeProxy.cs" company="Philipp Dolder">
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
    public class GroupShapeProxy : ICompositeNodeProxy
    {
        private readonly List<INodeProxy> children = new List<INodeProxy>();

        public GroupShapeProxy(string groupShapeName)
        {
            this.Name = groupShapeName;
        }

        public string Name { get; private set; }

        public IReadOnlyList<INodeProxy> Children
        {
            get { return this.children; }
        }

        public void Accept(NodeVisitor visitor)
        {
            visitor.VisitGroupShapeStart(this);

            foreach (INodeProxy child in this.children)
            {
                child.Accept(visitor);
            }

            visitor.VisitGroupShapeEnd(this);
        }

        public void Add(INodeProxy node)
        {
            this.children.Add(node);
        }
    }
}
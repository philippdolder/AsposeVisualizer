// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RowProxy.cs" company="Philipp Dolder">
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
    public class RowProxy : ICompositeNodeProxy
    {
        private readonly List<CellProxy> cells = new List<CellProxy>();

        public IReadOnlyList<CellProxy> Cells
        {
            get { return this.cells; }
        }

        public void Add(INodeProxy node)
        {
            this.cells.Add((CellProxy)node);
        }

        public virtual void Accept(NodeVisitor visitor)
        {
            visitor.VisitRowStart(this);

            foreach (CellProxy cell in this.cells)
            {
                cell.Accept(visitor);
            }

            visitor.VisitRowEnd(this);
        }
    }
}
// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeVisualizerObjectSource.cs" company="Philipp Dolder">
//   Copyright (c) 2013
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
    using System.IO;
    using Aspose.Words;
    using Microsoft.VisualStudio.DebuggerVisualizers;

    public class AsposeVisualizerObjectSource : VisualizerObjectSource
    {
        public override void GetData(object target, Stream outgoingData)
        {
            var writer = new StreamWriter(outgoingData);
            try
            {
                var root = (Node)target;
                var visitor = new XmlStructureDocumentVisitor();
                root.Accept(visitor);

                writer.Write(visitor.AsXml);
            }
            catch (InvalidCastException e)
            {
                string message = string.Concat(
                    "It seems the version of Aspose.Words you are debugging and the installed Aspose.Words Debugger Visualizer don't match.",
                    "\r\nPlease ensure both versions match.\r\n\r\n",
                    e.Message);

                writer.Write(message);
            }
            catch (Exception e)
            {
                writer.Write(e.Message);
            }

            writer.Flush();
        }
    }
}
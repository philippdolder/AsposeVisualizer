// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeDialogDebuggerVisualizer.cs" company="Philipp Dolder">
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
using Aspose.Words;
using AsposeVisualizer;

[assembly: System.Diagnostics.DebuggerVisualizer(
typeof(AsposeDialogDebuggerVisualizer),
typeof(AsposeVisualizerObjectSource),
Target = typeof(Node),
Description = "Aspose.Words Document Visualizer")]

namespace AsposeVisualizer
{
    using System.IO;
    using Microsoft.VisualStudio.DebuggerVisualizers;

    public class AsposeDialogDebuggerVisualizer : DialogDebuggerVisualizer
    {
        protected override void Show(IDialogVisualizerService windowService, IVisualizerObjectProvider objectProvider)
        {
            var msg = new StreamReader(objectProvider.GetData()).ReadToEnd();

            var form = new AsposeVisualizerForm();
            form.SetText(msg);

            windowService.ShowDialog(form);
        }
    }
}
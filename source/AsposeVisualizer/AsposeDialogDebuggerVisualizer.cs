// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeDialogDebuggerVisualizer.cs" company="Philipp Dolder">
//   Copyright (c) 2013-2014
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

using AsposeVisualizer;

[assembly: System.Diagnostics.DebuggerVisualizer(
typeof(AsposeDialogDebuggerVisualizer),
typeof(AsposeVisualizerObjectSource),
Target = typeof(Aspose.Words.Node),
Description = "Aspose.Words Document Visualizer")]

namespace AsposeVisualizer
{
    using Microsoft.VisualStudio.DebuggerVisualizers;

    public class AsposeDialogDebuggerVisualizer : DialogDebuggerVisualizer
    {
        public static void TestShowVisualizer(object node)
        {
            var host =
                new VisualizerDevelopmentHost(node, typeof(AsposeDialogDebuggerVisualizer), typeof(AsposeVisualizerObjectSource));
            host.ShowVisualizer();
        }

        protected override void Show(IDialogVisualizerService windowService, IVisualizerObjectProvider objectProvider)
        {
            var nodeProxy = (INodeProxy)objectProvider.GetObject();

            var window = new VisualizerWindow { DataContext = new VisualizerViewModel(nodeProxy) };

            window.ShowDialog();
        }
    }
}
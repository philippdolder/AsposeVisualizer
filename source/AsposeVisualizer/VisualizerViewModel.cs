// --------------------------------------------------------------------------------------------------------------------
// <copyright file="VisualizerViewModel.cs" company="Philipp Dolder">
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
    using System.ComponentModel;
    using System.Windows;
    using System.Windows.Input;

    public class VisualizerViewModel : INotifyPropertyChanged
    {
        private readonly INodeProxy documentProxy;
        private bool includeFormatting;
        private bool includeImages;

        public VisualizerViewModel(INodeProxy documentProxy)
        {
            this.documentProxy = documentProxy;
            this.CopyCommand = new RelayCommand(this.CopyToClipboard);
        }

        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public ICommand CopyCommand { get; private set; }

        public string Xml
        {
            get
            {
                return this.CreateXml();
            }
        }

        public bool IncludeFormatting
        {
            get
            {
                return this.includeFormatting;
            }

            set
            {
                this.includeFormatting = value;
                this.PropertyChanged(this, new PropertyChangedEventArgs("IncludeFormatting"));
                this.PropertyChanged(this, new PropertyChangedEventArgs("Xml"));
            }
        }

        public bool IncludeImages
        {
            get
            {
                return this.includeImages;
            }

            set
            {
                this.includeImages = value;
                this.PropertyChanged(this, new PropertyChangedEventArgs("IncludeImages"));
                this.PropertyChanged(this, new PropertyChangedEventArgs("Xml"));
            }
        }

        private void CopyToClipboard()
        {
            Clipboard.SetText(this.Xml);
        }

        private string CreateXml()
        {
            var visitor = new XmlStructureNodeVisitor(new XmlStructureDisplayOptions(this.IncludeFormatting, this.IncludeImages));
            this.documentProxy.Accept(visitor);

            return visitor.AsXml;
        }
    }
}
AsposeVisualizer
================

VS Debugger Visualizer for Aspose.Words.

This is (yet) a simple Visual Studio Debugger Visualizer that allows you to see the structure of all Aspose.Words.Node in an XML format to allow easier understanding of what your document looks like while debugging. Please raise a github [issue](https://github.com/philippdolder/AsposeVisualizer/issues) if you are missing anything.

Currently there is a zip file release available for current versions of Aspose.Words (downloadable from [Aspose](http://www.aspose.com/community/files/51/.net-components/aspose.words-for-.net/default.aspx)).


Installation instructions
-------------------------
* Download the AsposeVisualizer release for your Aspose.Words version from [here](https://github.com/philippdolder/AsposeVisualizer/releases). If your version is missing, contact me on twitter ([@philippdolder](https://twitter.com/philippdolder)) or raise an [issue](https://github.com/philippdolder/AsposeVisualizer/issues)
* Unblock and unzip to a dedicated folder
* Run install.ps1 (will need internet connection to download Aspose.Words from NuGet feed)


Known issues
------------
* Sometimes you get a "Unhandled exception has occurred [...]. Object is in a zombie state"
** This happens before the AsposeVisualizer comes into play, so unfortunately I have no chance to avoid this
** It seems to being caused by slow serialization
** Just click "Continue" and try again, it usually works perfect after the first zombie exception


Important notice
----------------
There is in no way a link between this project and Aspose Ltd., the company behind Aspose.Words. You are responsible yourself to have the appropriate license for Aspose.Words to use this product.

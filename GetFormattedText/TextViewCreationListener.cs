using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace GetFormattedText
{
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal class TextViewCreationListener : IWpfTextViewCreationListener
    {
        [Import]
        IClassificationFormatMapService _classificationFormatMapService = null;

        [Import]
        IViewTagAggregatorFactoryService _tagAggregatorFactory = null;

        public void TextViewCreated(IWpfTextView textView)
        {
            new TextViewChangeFormatter(textView, _tagAggregatorFactory, _classificationFormatMapService);
        }
    }
}

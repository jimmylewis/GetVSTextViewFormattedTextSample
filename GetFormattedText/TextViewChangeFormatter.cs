using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Text.Tagging;

namespace GetFormattedText
{
    class TextViewChangeFormatter
    {
        private ITextView _textView;
        private ITagAggregator<IClassificationTag> _classificationAggregator;
        private IClassificationFormatMap _classificationFormatMap;

        public TextViewChangeFormatter(ITextView textView, IViewTagAggregatorFactoryService tagAggregatorFactory, IClassificationFormatMapService classificationFormatMapService)
        {
            if (!textView.IsClosed)
            {
                _textView = textView;
                _classificationAggregator = tagAggregatorFactory.CreateTagAggregator<IClassificationTag>(_textView);
                _classificationFormatMap = classificationFormatMapService.GetClassificationFormatMap(textView);

                _textView.Closed += OnTextViewClosed;
                textView.TextBuffer.PostChanged += OnTextBufferPostChanged;
            }
        }

        private void OnTextBufferPostChanged(object sender, EventArgs e)
        {
            if (sender is ITextBuffer textBuffer)
            {
                string rtfString = GetFormattedBufferText(textBuffer);
                if (rtfString.Length > 0)
                {
                    MessageBoxResult result = MessageBox.Show("Copy text to clipboard?", "", MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.Yes)
                    {
                        Clipboard.SetText(rtfString, TextDataFormat.Rtf);
                    }
                }
            }
        }

        private string GetFormattedBufferText(ITextBuffer textBuffer)
        {
            var wholeBufferSpan = new SnapshotSpan(textBuffer.CurrentSnapshot, 0, textBuffer.CurrentSnapshot.Length);
            var textSpans = GetTextSpansWithFormatting(textBuffer);
            RTFStringBuilder sb = new RTFStringBuilder();

            int currentPos = 0;
            var formattedSpanEnumerator = textSpans.GetEnumerator();
            while (currentPos < wholeBufferSpan.Length && formattedSpanEnumerator.MoveNext())
            {
                var spanToFormat = formattedSpanEnumerator.Current;
                if (currentPos < spanToFormat.Span.Start)
                {
                    int unformattedLength = spanToFormat.Span.Start - currentPos;
                    SnapshotSpan unformattedSpan = new SnapshotSpan(textBuffer.CurrentSnapshot, currentPos, unformattedLength);
                    sb.AppendText(unformattedSpan.GetText(), System.Drawing.Color.Black);
                }

                System.Drawing.Color textColor = GetTextColor(spanToFormat.Formatting.ForegroundBrush);
                sb.AppendText(spanToFormat.Span.GetText(), textColor);

                currentPos = spanToFormat.Span.End;
            }

            if (currentPos < wholeBufferSpan.Length)
            {
                // append any remaining unformatted text
                SnapshotSpan unformattedSpan = new SnapshotSpan(textBuffer.CurrentSnapshot, currentPos, wholeBufferSpan.Length - currentPos);
                sb.AppendText(unformattedSpan.GetText(), System.Drawing.Color.Black);
            }

            return sb.ToString();
        }

        private void OnTextViewClosed(object sender, EventArgs e)
        {
            _textView.Closed -= OnTextViewClosed;
            _textView.TextBuffer.PostChanged -= OnTextBufferPostChanged;

            _classificationAggregator = null;
            _classificationFormatMap = null;
        }

        private System.Drawing.Color GetTextColor(Brush foregroundBrush)
        {
            System.Drawing.Color result = System.Drawing.Color.Black;

            if (foregroundBrush is SolidColorBrush textBrush)
            {
                result = System.Drawing.Color.FromArgb(textBrush.Color.A, textBrush.Color.R, textBrush.Color.G, textBrush.Color.B);
            }

            return result;
        }

        private IEnumerable<(SnapshotSpan Span, TextFormattingRunProperties Formatting)> GetTextSpansWithFormatting(ITextBuffer textBuffer)
        {
            var wholeBufferSpan = new SnapshotSpan(textBuffer.CurrentSnapshot, 0, textBuffer.CurrentSnapshot.Length);
            var tags = _classificationAggregator.GetTags(wholeBufferSpan);

            foreach (var tag in tags)
            {
                TextFormattingRunProperties formattingProperties = _classificationFormatMap.GetTextProperties(tag.Tag.ClassificationType);
                int startPoint = tag.Span.Start.GetPoint(textBuffer, PositionAffinity.Successor).Value.Position;
                Span textSpan = new Span(
                    startPoint,
                    tag.Span.End.GetPoint(textBuffer, PositionAffinity.Predecessor).Value.Position - startPoint);
                yield return (new SnapshotSpan(textBuffer.CurrentSnapshot, textSpan), formattingProperties);
            }
        }

    }
}

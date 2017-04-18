using System.Drawing;
using System.Windows.Forms;

namespace GetFormattedText
{
    class RTFStringBuilder
    {
        RichTextBox rtb = new RichTextBox();

        public void AppendText(string text, Color textColor)
        {
            rtb.SelectionColor = textColor;
            rtb.AppendText(text);
            
        }

        public override string ToString()
        {
            rtb.SelectAll();
            string result = rtb.SelectedRtf;
            rtb.Select(rtb.TextLength, 0);
            return result;
        }
    }
}

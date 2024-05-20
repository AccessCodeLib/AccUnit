using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public partial class TestResultDetailView : Window
    {
        public TestResultDetailView(TestResultViewModel dataContext)
        {
            InitializeComponent();
            DataContext = dataContext;
            richTextBox.Document = dataContext.HighlightedText;
            AdjustRichTextBoxWidth();
        }

        private void AdjustRichTextBoxWidth()
        {
            // Create a TextBlock to measure the width of the text
            TextBlock textBlock = new TextBlock
            {
                TextWrapping = TextWrapping.NoWrap, // Ensure no wrapping for accurate width
                FontFamily = richTextBox.FontFamily,
                FontSize = richTextBox.FontSize,
                FontStyle = richTextBox.FontStyle,
                FontWeight = richTextBox.FontWeight
            };

            double maxWidth = 0;

            // Measure the width of each line in the RichTextBox
            TextPointer lineStart = richTextBox.Document.ContentStart;
            while (lineStart != null && lineStart.CompareTo(richTextBox.Document.ContentEnd) < 0)
            {
                TextPointer lineEnd = lineStart.GetLineStartPosition(1);
                if (lineEnd == null)
                {
                    lineEnd = richTextBox.Document.ContentEnd;
                }

                string lineText = new TextRange(lineStart, lineEnd).Text.TrimEnd();
                textBlock.Text = lineText;

                textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                if (textBlock.DesiredSize.Width > maxWidth)
                {
                    maxWidth = textBlock.DesiredSize.Width;
                }

                lineStart = lineEnd.GetNextInsertionPosition(LogicalDirection.Forward);
            }

            double padding = 50;

            richTextBox.Width = maxWidth + padding;
        }
    }
}

using System.Diagnostics;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class SimpleLinkLabel : LinkLabel
    {
        public override string Text
        {
            get { return base.Text; }
            set
            {
                base.Text = value;
                var url = base.Text;
                Links.Clear();
                Links.Add(0, url.Length, url);
            }
        }

        protected override void OnLinkClicked(LinkLabelLinkClickedEventArgs e)
        {
            base.OnLinkClicked(e);
            var psi = new ProcessStartInfo {
                                               UseShellExecute = true,
                                               FileName = e.Link.LinkData.ToString()
                                           };
            Process.Start(psi);
        }
    }
}
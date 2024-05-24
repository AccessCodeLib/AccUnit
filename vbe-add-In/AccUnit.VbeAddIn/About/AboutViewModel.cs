using AccessCodeLib.AccUnit.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Input;

namespace AccessCodeLib.AccUnit.VbeAddIn.About
{
    public class AboutViewModel
    {
        public ICommand NavigateCommand { get; }

        public string AddInVersion => AddInManager.FileVersion;
        public string FrameworkVersion => AccUnitInfo.FileVersion;
        public string AddInCopyright => AddInManager.Copyright;
        public string FrameworkCopyright => AccUnitInfo.Copyright;
        public string Copyright => AddInCopyright.CompareTo(FrameworkCopyright) >= 0 
                                    ? AddInManager.Copyright : FrameworkCopyright;


        public AboutViewModel()
        {
            NavigateCommand = new RelayCommand<string>(Navigate);
            Contributors = new List<Contributor>
            {
                new Contributor("Josef Pötzl"),
                new Contributor("Paul Rohorzka"),
                new Contributor("Sten Schmidt")
            };
        }

        private void Navigate(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
            }
        }

        public IEnumerable<Contributor> Contributors { get; }

    }

    public class Contributor
    {
        public Contributor(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }

}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookDataBackup
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static SettingsHandler Settings { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            Settings = new SettingsHandler(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\OutlookDataBackup");

            base.OnStartup(e);
        }
    }
}

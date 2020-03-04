using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using Microsoft.Graph;
using Microsoft.Win32;

namespace OutlookDataBackup
{
    /// <summary>
    /// Interaction logic for FolderBrowserWindow.xaml
    /// </summary>
    public partial class FolderBrowserWindow : Window
    {
        internal GraphServiceClient graphClient;
        private readonly object dummyNode = null;
        internal string selectedPath = "";

        public FolderBrowserWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (foldersTreeView.SelectedItem == null)
            {
                MessageBox.Show("You haven't selected any destination path; please choose one", "No selection",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                return;
            }

            selectedPath = foldersTreeView.SelectedItem == foldersTreeView.Items[0] ? "/" : ((Tuple<string, string>)((TreeViewItem)foldersTreeView.SelectedItem).Tag).Item2;           

            this.graphClient = null;
            DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.graphClient = null;
            DialogResult = false;
            this.Close();
        }

        private void FolderBrowserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            var item = new TreeViewItem
            {
                Header = "OneDrive",
                Tag = new Tuple<string, string>("root", ""),
                FontWeight = FontWeights.Normal
            };

            item.Items.Add(dummyNode);
            item.Expanded += Folder_Expanded;

            foldersTreeView.Items.Add(item);
        }

        private async void Folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == dummyNode)
            {
                if (graphClient != null)
                {
                    item.Items.Clear();

                    if (item.Tag is Tuple<string, string> tag)
                    {
                        var children = item.Tag.ToString() == "root"
                            ? await graphClient?.Me.Drive.Root.Children.Request().GetAsync()
                            : await graphClient?.Me.Drive.Items[tag.Item1].Children.Request().GetAsync();

                        foreach (var subFolder in children.Where(x => x.Folder != null))
                        {
                            var subItem = new TreeViewItem
                            {
                                Header = subFolder.Name,
                                Tag = new Tuple<string, string>(subFolder.Id, tag.Item2 + "/" + subFolder.Name),
                                FontWeight = FontWeights.Normal
                            };

                            subItem.Items.Add(dummyNode);
                            subItem.Expanded += Folder_Expanded;

                            item.Items.Add(subItem);
                        }
                    }
                }
            }
        }

        private void FoldersTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = (TreeViewItem)e.NewValue;

            currentFolderTextBlock.Text = "Selected folder: " + item.Header;
        }
    }
}

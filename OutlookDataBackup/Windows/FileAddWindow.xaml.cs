using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Graph;
using Microsoft.Win32;
using MahApps.Metro.Controls;

namespace OutlookDataBackup
{
    /// <summary>
    /// Interaction logic for FileAddWindow.xaml
    /// </summary>
    public partial class FileAddWindow : Window
    {
        internal List<string> Files;

        public FileAddWindow()
        {
            InitializeComponent();
            Files = new List<string>();
        }

        /// <summary>
        /// Reset UI after a long async task
        /// </summary>
        /// <returns></returns>
        private async Task ResetUi()
        {
            await Dispatcher.InvokeAsync(() =>
            {
                okButton.IsEnabled = true;
                cancelButton.IsEnabled = true;
                checkProgressRing.IsActive = false;
                checkTextBlock.Visibility = Visibility.Collapsed;
            });
        }

        private void SourceChooseButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a new OpenFileDialog to pick the file
            var openFileDialog = new OpenFileDialog
            {
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "pst file(*.pst)|*.pst|ost file(*.ost)|*.ost|All Outlook Files(*.pst;*.ost)|*.pst;*.ost",
                FilterIndex = 3,
                Multiselect = true,
                ValidateNames = true
            };

            // Show the OpenFileDialog and get the selected file(s)
            var result = openFileDialog.ShowDialog();

            // Check that the user selected some file
            if (!result.HasValue || !result.Value) return;

            // Set the result
            Files = openFileDialog.FileNames.ToList();
            filesTextBox.Text = string.Join(", ", openFileDialog.FileNames);
        }

        private void DestinationChooseButton_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowserWindow = new FolderBrowserWindow
            {
                Owner = this,
                graphClient = (Owner as MainWindow)?.graphClient
            };

            var result = folderBrowserWindow.ShowDialog();

            // Check that the user selected some file
            if (!result.HasValue || !result.Value) return;

            // Check that the user selected some file
            if (string.IsNullOrEmpty(folderBrowserWindow.selectedPath)) return;

            destTextBox.Text = folderBrowserWindow.selectedPath;
        }

        private async void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Check that the user has selected some files
            if (string.IsNullOrEmpty(filesTextBox.Text))
            {
                MessageBox.Show("Please choose some files to backup.", "Enter a destination",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                return;
            }

            // Check that the TextBox is not empty
            if (string.IsNullOrEmpty(destTextBox.Text))
            {
                MessageBox.Show("Please choose a destination folder.", "Enter a destination",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                return;
            }

            try
            {
                var graphClient = (Owner as MainWindow)?.graphClient;
                if (graphClient != null)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        okButton.IsEnabled = false;
                        cancelButton.IsEnabled = false;
                        checkProgressRing.IsActive = true;
                        checkTextBlock.Visibility = Visibility.Visible;
                    });

                    // Check if the folder exists
                    if (destTextBox.Text.Equals("/"))
                    {
                        await graphClient?.Me.Drive.Root.Request().GetAsync();
                    }
                    else
                    {
                        await graphClient?.Me.Drive.Root.ItemWithPath(destTextBox.Text).Request().GetAsync();
                    }
                }
            }
            catch (Exception exception)
            {
                await ResetUi();

                var message = !(exception is ServiceException oneDriveException) ? exception.Message : $"{oneDriveException.Error.Message}";

                MessageBox.Show("Probably the specified destination doesn't exist."+
                                "\r\nNB: remember that the path must start with '/' and must not end with the slash.\r\n\r\nError: " + message,
                    "Path error", MessageBoxButton.OK, MessageBoxImage.Error);

                return;
            }

            // Reset UI, set DialogResult and close
            await ResetUi();
            DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Set DialogResult and close
            DialogResult = false;
            this.Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!DialogResult.HasValue) DialogResult = false;
        }
    }
}

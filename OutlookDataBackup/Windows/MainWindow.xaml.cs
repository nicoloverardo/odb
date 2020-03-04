using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Graph;
using Squirrel;
using MahApps.Metro.Controls;

namespace OutlookDataBackup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal GraphServiceClient graphClient;
        private const int UploadChunkSize = 5 * 1024 * 1024; // 5 MB
        private PstFile currentPst;
        private CancellationTokenSource cts;

        private NetworkTraffic trafficMonitor;
        private System.Timers.Timer bandwidthCalcTimer;
        private float lastAmountOfBytesSent;

        public MainWindow()
        {
            InitializeComponent();
            currentPst = new PstFile("", "", "", 1);

            bandwidthCalcTimer = new System.Timers.Timer(1000);
            bandwidthCalcTimer.Elapsed += BandwidthCalcTimer_Elapsed;
        }

        #region Login/Logout
        private async void LoginMenuItem_Click(object sender, RoutedEventArgs e)
        {
            cts = new CancellationTokenSource();
            var token = cts.Token;

            // Login and get the name of the user
            var name = await Task.Run(async () =>
            {
                try
                {
                    // Get the Microsoft Graph client
                    graphClient = await Task.Run(() => AuthenticationHelper.GetAuthenticatedClient(), token);

                    // Enable the cancel button
                    await Dispatcher.InvokeAsync(() =>
                    {
                        cancelButton.IsEnabled = true;
                    });

                    return "Welcome, " + (await graphClient.Me.Request().GetAsync(token)).DisplayName + "!";
                }
                catch (OperationCanceledException)
                {
                    MessageBox.Show("Login cancelled", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return "";
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, "Login error", MessageBoxButton.OK, MessageBoxImage.Error);
                    graphClient = null;

                    return "";
                }
            }, token);

            if (string.IsNullOrEmpty(name)) return;

            // Set the Ui
            await Dispatcher.InvokeAsync(() =>
            {
                loginMenuItem.IsEnabled = false;
                logoutMenuItem.IsEnabled = true;
                welcomeTextBlock.Visibility = Visibility.Visible;
                welcomeTextBlock.Text = name;
                //welcomeTextBlock.Click += LogoutMenuItem_Click;
                FileMenuItem.IsSubmenuOpen = false;
                addFileButton.IsEnabled = true;
                cancelButton.IsEnabled = false;
            });
        }

        private async void LogoutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Check if the user has logged in
            if (graphClient != null)
            {
                // Prompt the user to log out
                var result = MessageBox.Show("Do you really want to log out?", "Logout", MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                // Return if the user doesn't want to log out
                if (result != MessageBoxResult.Yes) return;

                // Log out
                await AuthenticationHelper.SignOut();
            }

            MessageBox.Show("Logged out successfully", "Logout", MessageBoxButton.OK, MessageBoxImage.Information);

            await Dispatcher.InvokeAsync(() =>
            {
                loginMenuItem.IsEnabled = true;
                logoutMenuItem.IsEnabled = false;
                addFileButton.IsEnabled = false;
                startBackupButton.IsEnabled = false;
                cancelButton.IsEnabled = false;
                welcomeTextBlock.Visibility = Visibility.Collapsed;
                welcomeTextBlock.Text = "";
                //welcomeTextBlock.Click += LoginMenuItem_Click;
            });
        }
        #endregion

        #region Add/Remove files
        private void AddFileButton_Click(object sender, RoutedEventArgs e)
        {
            var files = new List<PstFile>();

            // We use this to track which file eventually raises any error
            string currentFileAdded = "";

            try
            {
                var addFileButtonWindow = new FileAddWindow
                {
                    Owner = this
                };
                var result = addFileButtonWindow.ShowDialog();

                if (!result.HasValue) return;
                if (!result.Value) return;

                foreach (var file in addFileButtonWindow.Files)
                {
                    currentFileAdded = file;

                    var fileInfo = new FileInfo(file);
                    files.Add(new PstFile(fileInfo.Name, fileInfo.FullName, addFileButtonWindow.destTextBox.Text, fileInfo.Length));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error while adding file " + currentFileAdded + " to the list.\r\n\r\nError: " + ex.Message,
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                return;
            }

            if (filesListView.ItemsSource != null)
            {
                var newList = new List<PstFile>(filesListView.ItemsSource.Cast<PstFile>());
                newList.AddRange(files);
                filesListView.ItemsSource = newList;
            }
            else
            {
                filesListView.ItemsSource = files;
            }

            if (filesListView.Items.Count > 0)
            {
                removeFilesButton.IsEnabled = true;
                startBackupButton.IsEnabled = true;
            }
        }

        private async void RemoveFilesButton_Click(object sender, RoutedEventArgs e)
        {
            if (filesListView.SelectedItem == null) return;

            var selected = filesListView.SelectedItems.Cast<PstFile>().ToList();
            var allItems = filesListView.ItemsSource as List<PstFile>;

            cts = new CancellationTokenSource();
            var token = cts.Token;

            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    foreach (var item in selected)
                    {
                        token.ThrowIfCancellationRequested();
                        if (item != currentPst) allItems?.Remove(item);
                    }
                }, DispatcherPriority.Normal, token);

                filesListView.ItemsSource = null;
                filesListView.ItemsSource = allItems;
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Task cancelled", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                if (filesListView.Items.Count == 0 || (filesListView.Items.Count == 1 && filesListView.Items[0] == currentPst))
                {
                    removeFilesButton.IsEnabled = false;
                    startBackupButton.IsEnabled = false;
                }

                cts = null;
            }
        }
        #endregion    

        #region Updates
        private async void UpdatesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Check if is Squirrel app
            if (!IsSquirrelInstalledApp())
            {
                MessageBox.Show("App is not installed using Squirrel", "Can't update", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                // In order to be able to use Dropbox as a repository for the setup files,
                // we have decomposed the update process as follows:
                // 1. We have created a public Dropbox folder that contains an index file named "updates.txt"
                //    that contains the link to each file needed (see the Squirrel documentation on GitHub for more info).
                //    Moreover, in the first line, the index file has the last version available online, so that
                //    it allows us to check for updates remotely. Finally, "?raw=1" at the end of each link allows us
                //    to directly download the file.
                // 2. We download this index file and parse it.
                // 3. We check the if the user has the latest version. If he hasn't, then we proceed.
                // 4. We download then each necessary file. Note that "DELTA" .nuget file may not be present, so we check for that.
                // 5. Once we have downloaded all files into a temp folder, we run the Squirrel update manager
                //    and let it do its magic.

                // So, initialize a WebClient in order to download the required files.
                using (var client = new WebClient())
                {
                    // Download the index file
                    var indexFile = await client.DownloadStringTaskAsync(
                        "https://www.dropbox.com/sh/e0rw5dfuoh9jgjt/AAA1wdaCiqMWRYsoqY2cZlNra/updates.txt?raw=1");

                    // Parse it
                    var indexes = Regex.Split(indexFile, "\r\n|\r|\n").ToList();

                    // Get the last version available
                    Version.TryParse(indexes[0].Substring(8) + ".0", out var version);
                    if (version == null)
                    {
                        MessageBox.Show("Cannot parse the new version", "Error", MessageBoxButton.OK,
                            MessageBoxImage.Error);

                        return;
                    }

                    // Check it with the installed version
                    if (version > Assembly.GetExecutingAssembly().GetName().Version)
                    {
                        var message = new StringBuilder().AppendLine($"A new version is available ({version})?").
                            AppendLine("If you choose to update, changes will not take affect until the app is restarted.").
                            AppendLine("Would you like to download and install it?").
                            ToString();

                        var result = MessageBox.Show(message, "App Update", MessageBoxButton.YesNo, MessageBoxImage.Information);
                        if (result != MessageBoxResult.Yes) return;

                        // Delete old temp folder
                        if (System.IO.Directory.Exists(Path.GetTempPath() + "OutlookDataBackup"))
                            System.IO.Directory.Delete(Path.GetTempPath() + "OutlookDataBackup", true);

                        // And recreate it
                        System.IO.Directory.CreateDirectory(Path.GetTempPath() + "OutlookDataBackup");

                        // Download the RELEASES file
                        await client.DownloadFileTaskAsync(indexes[1].Replace("releases=", ""),
                            Path.GetTempPath() + "OutlookDataBackup\\RELEASES");

                        // Download the FULL .nuget file
                        await client.DownloadFileTaskAsync(indexes[2].Replace("full=", ""),
                            Path.GetTempPath() + "OutlookDataBackup\\" +
                            indexes[2].Remove(indexes[2].IndexOf("?raw=1", StringComparison.Ordinal))
                                .Substring(indexes[2].LastIndexOf("/", StringComparison.Ordinal) + 1));

                        // Download the DELTA .nuget file if available
                        if (indexes[3] != "delta=")
                        {
                            await client.DownloadFileTaskAsync(indexes[3].Replace("delta=", ""),
                                Path.GetTempPath() + "OutlookDataBackup\\" +
                                indexes[3].Remove(indexes[3].IndexOf("?raw=1", StringComparison.Ordinal))
                                    .Substring(indexes[3].LastIndexOf("/", StringComparison.Ordinal) + 1));
                        }

                        // Download the Setup.exe file
                        await client.DownloadFileTaskAsync(indexes[4].Replace("setupexe=", ""),
                            Path.GetTempPath() + "OutlookDataBackup\\Setup.exe");

                        // And finally, the Setup.msi file
                        await client.DownloadFileTaskAsync(indexes[5].Replace("setupmsi=", ""),
                            Path.GetTempPath() + "OutlookDataBackup\\Setup.msi");

                        // Now that we have all the necessary files, initialize the update manager
                        using (var manager = new UpdateManager(Path.GetTempPath() + "OutlookDataBackup"))
                        {
                            // And then launch the update process.
                            var updateResult = await manager.UpdateApp();

                            MessageBox.Show($"Download complete. Please restart to install version {updateResult.Version}.", "Updating",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("You have the latest version.", "No updates", MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while updating:\r\n\r\n" + ex.Message, "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Checks if the app is installed using <see cref="Squirrel"/>
        /// </summary>
        /// <returns>True if the app is installed using Squirrel</returns>
        private static bool IsSquirrelInstalledApp()
        {
            try
            {
                var updateDotExe = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) ?? throw new Exception(),
                    "..", "Update.exe");

                return System.IO.File.Exists(updateDotExe);
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Upload
        private async void StartBackupButton_Click(object sender, RoutedEventArgs e)
        {
            cts = new CancellationTokenSource();
            var token = cts.Token;

            // Setup UI
            await Dispatcher.InvokeAsync(() =>
            {
                itemProgressBar.Visibility = Visibility.Visible;
                zipProgressBar.Visibility = Visibility.Visible;
                itemProgressTextBlock.Visibility = Visibility.Visible;
                zipProgressTextBlock.Visibility = Visibility.Visible;
                itemProgressBar.Value = 0;
                zipProgressBar.Value = 0;
                zipProgressTextBlock.Text = "";
                itemProgressTextBlock.Text = "";
                startBackupButton.IsEnabled = false;
                cancelButton.IsEnabled = true;
            });

            // Check if Outlook is running
            try
            {
                if (IsOutlookRunning())
                {
                    MessageBox.Show("Outlook is currently running. Please close any instance of Outlook and try again", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    await ResetUiAsync();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Couldn't check if Outlook is currently open.\r\n\r\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                await ResetUiAsync();
                return;
            }

            // Check if OneDrive has enough free space
            try
            {
                var hasSpace = await CheckOneDriveAvailableSpace(token);

                if (!hasSpace)
                {
                    MessageBox.Show("Not enough space on OneDrive!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    await ResetUiAsync();
                    return;
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Task cancelled", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Warning);

                await ResetUiAsync();

                return;
            }

            // Set progress to the UI
            await Dispatcher.InvokeAsync(() =>
            {
                zipProgressTextBlock.Text = ((List<PstFile>)filesListView.ItemsSource).All(x => !x.NeedsZip) ? "No file needs to be zipped" : "";
            });

            var zipOccurred = false;

            // Stores the zip parts of each file
            var zipParts = new Dictionary<PstFile, List<PstFile>>();

            // Zip
            try
            {
                await Task.Run(async () =>
                {
                    // Iterate through items
                    foreach (var item in ((List<PstFile>)filesListView.ItemsSource))
                    {
                        token.ThrowIfCancellationRequested();

                        // Signal necessary to avoid that the user removes the
                        // currently uploading item
                        currentPst = item;

                        // Set progress to the UI
                        await Dispatcher.InvokeAsync(() =>
                        {
                            zipProgressBar.Value = 0;
                            zipProgressTextBlock.Text = "Zipping " + item.Name + "... (0%)";
                        });

                        // If the item does not exceed OneDrive limit,
                        // there won't be zip parts. Therefore,
                        // add the item with an empty list of zip parts
                        // and proceed to the next item
                        if (!item.NeedsZip)
                        {
                            zipParts.Add(item, new List<PstFile>());
                            continue;
                        }

                        // Zip and return the zip parts
                        var zips = await ZipFile(item, "-v2g"); //TODO: set to dynamic split size or -v2g
                        if (zips != null)
                        {
                            // Add them to the dictionary
                            zipParts.Add(item, zips);
                            zipOccurred = true;
                        }
                        else
                        {
                            throw new Exception("Something went wrong while zipping. Please try again.");
                        }
                    }
                }, token);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Task cancelled", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Warning);
                await ResetUiAsync();
                currentPst = null;
                return;
            }
            catch (Exception exception)
            {
                if (exception.InnerException != null)
                {
                    MessageBox.Show(exception.Message + "\r\n\r\n" + exception.InnerException.Message, "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show(exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                await ResetUiAsync();
                currentPst = null;
                return;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                zipProgressBar.Value = 0;
                zipProgressTextBlock.Text = zipOccurred ? "Zipping done" : "No file needs to be zipped";
            });

            /* TODO: resume upload session
             * 0. Design and create the JSON config file.
             * 1. List all the files that should be added.
             * 1.1. Get the original file path and its destination.
             * 1.2. Get all the zip parts path, destination and hash codes.
             *      Hash code will allow to check if the file has been modified
             *      and to abort the session in that case.
             * 2. Save this list in the config file.
             * 2.2. If the user removes a file from the ListView during upload,
             *      also remove it from the list in the config file.
             * 3. Save also the default behaviour in case of conflict for an
             *    already existing file.
             * 4. When a zipPart or the original file is uploading,
             *    get the uploadUrl of the UploadSession and save it.
             * 5. After each zipPart upload is completed, then update the
             *    flag in the config file that indicates if it is currently
             *    uploading.
             * 6. In the case of restore, then ask the user if we need to restore.
             *    If the user says yes, the:
             * 6.1. Look at the config file, and re-add the file to be uploaded.
             * 6.1.1. Before adding, compare hash codes of each part.
             * 6.1.2. In case of mismatch, abort the restore and clear the config file.
             * 6.2. Re-create the upload session (get remaining bytes etc...).
             * 6.3. Update the UI and progress to match the server progress percentage.
             * 7. When the backup is done, clear the config file.
             */

            //string origFileSourcePath = zipParts.FirstOrDefault().Key.Path;
            //string origFileDestPath = zipParts.FirstOrDefault().Key.Destination;
            //if (zipParts.FirstOrDefault().Value.Count != 0)
            //{
            //    foreach (var part in zipParts.FirstOrDefault().Value)
            //    {
            //        string partFileSourcePath = part.Path;
            //        string partFileDestPath = part.Destination;
            //        string hash = new FileInfo(part.Path).GetHashCode().ToString();
            //    }
            //}

            // Dictionary that will contain the filename and its upload id
            var uploadIds = new Dictionary<string, string>();

            // Get conflict behaviour action
            var conflict = conflictComboBox.SelectedIndex == 0 ? "replace" : "rename";

            // Initialize file counter
            var counter = 1;

            // Setup to check upload speed 
            trafficMonitor = new NetworkTraffic(System.Diagnostics.Process.GetCurrentProcess().Id);
            bandwidthCalcTimer.Enabled = true;

            try
            {
                await Task.Run(async () =>
                {
                    // Iterate through items
                    foreach (var item in zipParts)
                    {
                        token.ThrowIfCancellationRequested();

                        // Signal necessary to avoid the user from removing the
                        // the item currently uploading
                        currentPst = item.Key;

                        // Check if the pst/ost file was zipped
                        if (item.Value.Count == 0)
                        {
                            KeyValuePair<string, string> result;
                            
                            // Upload with the correct method according to the file size
                            if (item.Key.Length > 10485760) // 10 MiB
                            {
                                // Run the upload
                                result = await UploadFile(item.Key, conflict, false, zipParts.Count, counter, token);
                            }
                            else
                            {
                                await Dispatcher.InvokeAsync(() =>
                                {
                                    itemProgressBar.Value = 0;
                                    itemProgressBar.IsIndeterminate = true;

                                    itemProgressTextBlock.Text = "Uploading " + item.Key.Name + "...";
                                });

                                var uploadPath = item.Key.Destination + "/" + item.Key.Name;
                                var uploadedItem = await graphClient.Me.Drive.Root.ItemWithPath(uploadPath).Content.Request()
                                .PutAsync<DriveItem>(new FileStream(item.Key.Path, FileMode.Open), token);

                                result = new KeyValuePair<string, string>(item.Key.Name, uploadedItem.Id);

                                await Dispatcher.InvokeAsync(() =>
                                {
                                    itemProgressBar.IsIndeterminate = false;
                                });

                                //throw new Exception("File is smaller than 10 MiB. Please upload directly from the OneDrive windows app or website.");
                            }

                            // Check if the item was uploaded and then add its id to the list
                            if (!string.IsNullOrEmpty(result.Key) && !string.IsNullOrEmpty(result.Value))
                            {
                                uploadIds.Add(result.Key, result.Value);
                            }
                        }
                        else
                        {
                            // Initialize part counter
                            var i = 1;

                            // Upload every zip part
                            foreach (var part in item.Value)
                            {
                                token.ThrowIfCancellationRequested();

                                // Run the upload
                                var result = await UploadFile(part, conflict, true, zipParts.Count, counter, cts.Token, item.Value.Count, i);

                                // Check if the item was uploaded and then add its id to the list
                                if (!string.IsNullOrEmpty(result.Key) && !string.IsNullOrEmpty(result.Value))
                                {
                                    uploadIds.Add(result.Key, result.Value);
                                }

                                // Increment the part counter
                                i++;
                            }
                        }

                        // Increment the file counter
                        counter++;
                    }
                }, token);

                // Show upload ids
                var builder = new StringBuilder("Uploaded items and corresponding ids:\r\n\r\n");
                foreach (var item in uploadIds)
                {
                    builder.AppendLine(item.Key + " - " + item.Value);
                }

                MessageBox.Show(builder.ToString(), "Upload complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Task cancelled", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Error uploading", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Reset UI
                await ResetUiAsync();
                currentPst = null;
            }
        }

        /// <summary>
        /// Uploads a PST/OST file to OneDrive
        /// </summary>
        /// <param name="item">The file to upload</param>
        /// <param name="conflict">Indicates what to do in case the file already exists</param>
        /// <param name="multipleParts">Indicates if the pst/ost file was zipped and therefore split</param>
        /// <param name="totalFiles">The total number of files to upload</param>
        /// <param name="token">The <see cref="CancellationToken"/> used to check if the task should be cancelled.</param>
        /// <param name="file">The current file number</param>
        /// <param name="totalParts">The total numbers of parts, if the file was split</param>
        /// <param name="part">The current part, if the file was split</param>
        /// <returns>A pair containing the file name with its upload id.</returns>
        private async Task<KeyValuePair<string, string>> UploadFile(PstFile item, string conflict, bool multipleParts, int totalFiles, int file,
            CancellationToken token, int totalParts = 0, int part = 0)
        {
            // Set UI for progress report
            await Dispatcher.InvokeAsync(() =>
            {
                itemProgressTextBlock.Text = multipleParts
                ? "Uploading file " + file + " of " + totalFiles + ": " + item.Name + " | part " + part + "/" + totalParts + "... (0%)"
                : "Uploading file " + file + " of " + totalFiles + ": " + item.Name + "... (0%)";

                itemProgressBar.Value = 0;
            });

            // Run the chunks file upload
            var result = multipleParts
                ? await ChunkUpload(item, conflict, true, totalFiles, file, token, totalParts, part)
                : await ChunkUpload(item, conflict, false, totalFiles, file, token);

            // Check if the item was uploaded and then return its id
            return result != null ? new KeyValuePair<string, string>(result.Name, result.Id) : new KeyValuePair<string, string>();
        }

        /// <summary>
        /// Uploads a file by splitting it in chunks.
        /// </summary>
        /// <param name="item">The <see cref="PstFile"/> to upload.</param>
        /// <param name="conflictBehaviour">Indicates what to do in case the file already exists</param>
        /// <param name="totalFiles">The total number of files to upload</param>
        /// <param name="file">The current file number</param>
        /// <param name="token">The <see cref="CancellationToken"/> used to check if the task should be cancelled.</param>
        /// <param name="multipleParts">Indicates if the pst/ost file was zipped and therefore split</param>
        /// <param name="totalParts">The total numbers of parts, if the file was split</param>
        /// <param name="part">The current part, if the file was split</param>
        /// <returns>The uploaded file as <see cref="DriveItem"/></returns>
        private async Task<DriveItem> ChunkUpload(PstFile item, string conflictBehaviour, bool multipleParts, int totalFiles, int file,
            CancellationToken token, int totalParts = 1, int part = 1)
        {
            // Setup to check upload speed       
            lastAmountOfBytesSent = 0;          

            // Create the stream with the file
            using (var stream = new FileStream(item.Path, FileMode.Open))
            {
                // Set the upload path
                var uploadPath = item.Destination + "/" + item.Name;

                // Initialize the chunks provider
                ChunkedUploadProvider provider = null;

                // Will store the uploaded item
                DriveItem itemResult = null;

                // Create upload session
                var uploadSession = await graphClient.Me.Drive.Root.ItemWithPath(uploadPath)
                    .CreateUploadSession(new DriveItemUploadableProperties()
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", conflictBehaviour }
                        }
                    })
                    .Request()
                    .PostAsync(token);

                // Get the chunks provider
                provider = new ChunkedUploadProvider(uploadSession, graphClient, stream, UploadChunkSize);

                // Setup the chunk request necessities
                var chunkRequests = provider.GetUploadChunkRequests();
                //var readBuffer = new byte[UploadChunkSize];

                // Initialize counters for progress
                var maximum = chunkRequests.Count();
                var i = 1;

                // Upload the chunks
                await Task.Run(async () =>
                {
                    foreach (var request in chunkRequests)
                    {
                        // Delete session and throw cancellation if requested
                        if (token.IsCancellationRequested)
                        {
                            if (provider != null) await provider.DeleteSession();
                            token.ThrowIfCancellationRequested();
                        }

                        // Send chunk request
                        //var result = await provider.GetChunkRequestResponseAsync(request, readBuffer, new List<Exception>());
                        var result = await provider.GetChunkRequestResponseAsync(request, new List<Exception>());

                        // Update the itemProgressBar and UI
                        var currentValue = Math.Round((i / (double)maximum) * 100);
                        await Dispatcher.InvokeAsync(() =>
                        {
                            itemProgressBar.Value = currentValue;

                            itemProgressTextBlock.Text = multipleParts
                                ? "Uploading file " + file + " of " + totalFiles + ": " + item.Name + " | part " +
                                  part + "/" + totalParts + "... (" + itemProgressBar.Value.ToString("0") + "%)"
                                : "Uploading file " + file + " of " + totalFiles + ": " + item.Name + "... (" +
                                  itemProgressBar.Value.ToString("0") + "%)";
                        });

                        // Increment counter
                        i++;

                        if (result.UploadSucceeded)
                        {
                            itemResult = result.ItemResponse;
                        }
                    }
                }, token);

                return itemResult;
            }
        }

        private async void BandwidthCalcTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            float currentAmountOfBytesSent = trafficMonitor.GetBytesSent();

            await Dispatcher.InvokeAsync(() =>
            {
                zipProgressTextBlock.Text = "Upload speed: " + Helpers.GetReadableSize((long)(currentAmountOfBytesSent - lastAmountOfBytesSent)) + "/s";
            });

            lastAmountOfBytesSent = currentAmountOfBytesSent;
        }

        /// <summary>
        /// Checks if Outlook is running
        /// </summary>
        /// <returns>True if Outlook is running</returns>
        private bool IsOutlookRunning()
        {
            var process = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
            return process.Length != 0;
        }
        #endregion

        #region Zip
        /// <summary>
        /// Delete previous archive parts if present
        /// </summary>
        /// <param name="item">The name of the pst that will be used to know which file to delete</param>
        /// <returns></returns>
        private static async Task DeleteTempSplit(PstFile item)
        {
            await Task.Run(() =>
            {
                foreach (var file in new DirectoryInfo(Path.GetTempPath()).GetFiles())
                {
                    if (file.Name.Contains(item.Name + ".7z"))
                    {
                        file.Delete();
                    }
                }
            });
        }

        /// <summary>
        /// Creates a 7z archive asynchronously splitting the pst/ost file.
        /// </summary>
        /// <param name="item">The file to add to the archive</param>
        /// <param name="splitSize">The size of each part of the archive</param>
        /// <returns>A <see cref="List"/> of the parts of the archive</returns>
        private async Task<List<PstFile>> ZipFile(PstFile item, string splitSize)
        {
            cts = new CancellationTokenSource();
            var token = cts.Token;

            // Check for free space
            var freeSpace = await GetAvailableFreeSpace(Path.GetPathRoot(Path.GetTempPath()), token);

            if (freeSpace == -1)
            {
                throw new Exception(
                    "Couldn't compute the available space on the disk where the temp folder is located");
            }

            if (freeSpace <= item.Length)
            {
                throw new Exception("Not enough space on the drive " + Path.GetPathRoot(Path.GetTempPath()) +
                                    ".\r\nPlease free up space.");
            }

            // Delete previous archive parts if present
            try
            {
                await DeleteTempSplit(item);
            }
            catch (Exception e)
            {
                throw new Exception(
                    "Couldn't delete the previous 7zip parts temporary files. You can try delete them manually and then launch the backup again." +
                    "They are located in the temp folder: " + Path.GetTempPath() + "\r\n\r\n" + e.Message, e);
            }

            // Create process
            using (var pProcess = new System.Diagnostics.Process())
            {
                // strCommand is path and file name of command to run
                try
                {
                    pProcess.StartInfo.FileName =
                        Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "7za.exe");
                }
                catch (Exception e)
                {
                    throw new Exception("Error while finding the 7zip exe: " + e.Message, e);
                }

                // strCommandParameters are parameters to pass to program
                pProcess.StartInfo.Arguments = @"a " + "\"" + Path.GetTempPath() + item.Name + ".7z\" \"" + item.Path + "\" -mx0 -bsp1 " + splitSize; //splitSize: -v2g
                pProcess.StartInfo.UseShellExecute = false;

                // Set error and output of program to be written to process output stream
                pProcess.StartInfo.RedirectStandardOutput = true;
                pProcess.StartInfo.RedirectStandardError = true;

                // Hide the window
                pProcess.StartInfo.CreateNoWindow = true;
                //pProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

                // Process program output
                pProcess.OutputDataReceived += async (sender, args) =>
                {
                    // Check if args brings the progress
                    if (args.Data == null) return;
                    if (!Regex.IsMatch(args.Data, @"(\d+)%")) return;

                    // Get the progress
                    var progress = Regex.Match(args.Data, @"(\d+)%").Value.Replace("%", "");

                    // Update the UI
                    await Dispatcher.InvokeAsync(() =>
                    {
                        zipProgressBar.Value = Convert.ToDouble(progress);
                        zipProgressTextBlock.Text = "Zipping " + item.Name + "... (" + progress + "%)";
                    });
                };

                string stdError;
                try
                {
                    // Start the process
                    pProcess.Start();

                    // Asynchronously read the standard output of the spawned process. 
                    // This raises OutputDataReceived events for each line of output.
                    pProcess.BeginOutputReadLine();

                    // Asynchronously read the standard error of the spawned process.  
                    stdError = await pProcess.StandardError.ReadToEndAsync();

                    // Wait for process to finish
                    await Task.Run(() => { pProcess.WaitForExit(); }, token);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception e)
                {
                    if (!pProcess.HasExited) pProcess.Kill();
                    throw new Exception("Error while zipping: " + e.Message, e);
                }

                // Probably this first part of the if clause is useless
                if (!pProcess.HasExited)
                {
                    var result = MessageBox.Show(
                        "The zipping process should have been done by this point, but it isn't." +
                        "Click 'Yes' if you like to wait another minute for it to complete or 'No' to kill it and try again.",
                        "Zipping", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        // Wait for process to finish
                        await Task.Run(() => { pProcess.WaitForExit(60000); }, token);
                    }
                    else
                    {
                        pProcess.Kill();
                        throw new Exception("Zipping stopped. Try again please.");
                    }
                }
                else
                {
                    // Gather and show exceptions if there were errors
                    if (pProcess.ExitCode != 0)
                    {
                        var message = new StringBuilder();
                        if (!string.IsNullOrEmpty(stdError))
                        {
                            message.AppendLine(stdError);
                        }

                        throw new Exception("Zipping finished with exit code = " + pProcess.ExitCode + ": " + message);
                    }

                    // Get the list of the split 7z files
                    try
                    {
                        var finalList = new List<PstFile>();
                        foreach (var file in new DirectoryInfo(Path.GetTempPath()).GetFiles())
                        {
                            if (file.Name.Contains(item.Name))
                            {
                                finalList.Add(new PstFile(file.Name, file.FullName, item.Destination, file.Length));
                            }
                        }

                        return finalList;
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Error while getting the split files: " + e.Message, e);
                    }
                }

                return null;
            }
        }
        #endregion

        #region FreeSpace
        /// <summary>
        /// Asynchronously gets the free space of the selected drive
        /// </summary>
        /// <param name="driveName">The drive to check for free space</param>
        /// <param name="token">The <see cref="CancellationToken"/> used to check if the task should be cancelled.</param>
        /// <returns>The available space in bytes</returns>
        /// <remarks>Code taken and edited from https://stackoverflow.com/a/6815482 </remarks>
        private static async Task<long> GetAvailableFreeSpace(string driveName, CancellationToken token = default)
        {
            long freeSpace = -1;

            try
            {
                await Task.Run(() =>
                {
                    foreach (var drive in DriveInfo.GetDrives())
                    {
                        token.ThrowIfCancellationRequested();
                        if (drive.IsReady && drive.Name == driveName)
                        {
                            freeSpace = drive.AvailableFreeSpace;
                        }
                    }
                }, token);
            }
            catch
            {
                return freeSpace;
            }

            return freeSpace;
        }

        /// <summary>
        /// Asynchronously checks if OneDrive has enough space to store all the files that will be uploaded.
        /// </summary>
        /// <param name="token">The <see cref="CancellationToken"/> used to check if the task should be cancelled.</param>
        /// <returns>True if it has enough space; otherwise false.</returns>
        private async Task<bool> CheckOneDriveAvailableSpace(CancellationToken token = default)
        {
            // Check if OneDrive has some free space
            try
            {
                var quotaRemaining = (await graphClient.Me.Drive.Request().GetAsync(token)).Quota.Remaining;
                if (!quotaRemaining.HasValue) return false;

                // Get the quota remaining space
                var freeSpace = quotaRemaining.Value;

                long totalSize = 0;
                await Task.Run(() =>
                {
                    // Compute the total space required
                    foreach (var item in ((List<PstFile>)filesListView.ItemsSource))
                    {
                        token.ThrowIfCancellationRequested();
                        totalSize += item.Length;
                    }
                }, token);

                return totalSize < freeSpace;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Version " + Assembly.GetExecutingAssembly().GetName().Version + "\r\n\r\n" +
                            "Icon made by phatplus from www.flaticon.com\r\n\r\n" +
                            "Microsoft Graph SDK for .NET\r\n----------------\r\n" + "Copyright 2016 Microsoft Corporation\r\n" + "Licensed under the 'MIT License':\r\n" + "https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/LICENSE.txt" + "\r\n\r\n" +
                            "7zip version 19.0\r\n----------------\r\n" + "Copyright (C) 1999-2019 Igor Pavlov.\r\n" + "Licensed under 'GNU LGPL':\r\nhttps://www.7-zip.org/license.txt" + "\r\n\r\n" +
                            "Squirrel.Windows\r\n----------------\r\n" + "Licensed under the 'MIT License':\r\nhttps://github.com/Squirrel/Squirrel.Windows/blob/master/COPYING",
                "About Outlook Data Backup", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (cts == null) return;

            var result = MessageBox.Show("Would you like to cancel the running task?", "Cancelling",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                cts.Cancel();
                cancelButton.IsEnabled = false;
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Not yet implemented", "ODB", MessageBoxButton.OK, MessageBoxImage.Warning);
            var settingsWindow = new SettingsWindow()
            {
                Owner = this
            };

            settingsWindow.Show();
        }

        /// <summary>
        /// Reset UI after a long async task
        /// </summary>
        /// <returns></returns>
        private async Task ResetUiAsync()
        {
            // Reset UI
            await Dispatcher.InvokeAsync(() =>
            {
                itemProgressBar.Visibility = Visibility.Collapsed;
                zipProgressBar.Visibility = Visibility.Collapsed;
                itemProgressTextBlock.Visibility = Visibility.Collapsed;
                zipProgressTextBlock.Visibility = Visibility.Collapsed;
                itemProgressBar.Value = 0;
                zipProgressBar.Value = 0;
                zipProgressTextBlock.Text = "";
                itemProgressTextBlock.Text = "";
                startBackupButton.IsEnabled = true;
                cancelButton.IsEnabled = false;
            });

            if (bandwidthCalcTimer != null)
            {
                bandwidthCalcTimer.Enabled = false;
            }
        }
    }
}

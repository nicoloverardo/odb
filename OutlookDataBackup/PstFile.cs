namespace OutlookDataBackup
{
    public class PstFile
    {
        /// <summary>
        /// Represents a PST or OST file.
        /// </summary>
        /// <param name="name">File name, including extension</param>
        /// <param name="path">File full local path</param>
        /// <param name="destination">File full online path (i.e: /Documents/Backup/myfile.pst)</param>
        /// <param name="length">File length as long.</param>
        public PstFile(string name, string path, string destination, long length)
        {
            Name = name;
            Path = path;
            Destination = destination;
            Length = length;
            Size = Helpers.GetReadableSize(Length);
            Progress = 0;
            NeedsZip = SizeExceedsLimit(Length);
        }

        /// <summary>
        /// Represents a PST or OST file.
        /// </summary>
        /// <param name="name">File name, including extension</param>
        /// <param name="path">File full local path</param>
        /// <param name="destination">File full online path (i.e: /Documents/Backup/myfile.pst)</param>
        /// <param name="length">File length as long.</param>
        /// <param name="progress">Indicates the current upload progress of the file</param>
        public PstFile(string name, string path, string destination, long length, double progress)
        {
            Name = name;
            Path = path;
            Destination = destination;
            Length = length;
            Size = Helpers.GetReadableSize(Length);
            Progress = progress;
            NeedsZip = SizeExceedsLimit(Length);
        }

        /// <summary>
        /// Represents a PST or OST file.
        /// </summary>
        /// <param name="name">File name, including extension</param>
        /// <param name="path">File full local path</param>
        /// <param name="destination">File full online path (i.e: /Documents/Backup/myfile.pst)</param>
        /// <param name="length">File length as long.</param>
        /// <param name="progress">Indicates the current upload progress of the file</param>
        /// <param name="hash">Hash value</param>
        public PstFile(string name, string path, string destination, long length, double progress, int hash)
        {
            Name = name;
            Path = path;
            Destination = destination;
            Length = length;
            Size = Helpers.GetReadableSize(Length);
            Progress = progress;
            NeedsZip = SizeExceedsLimit(Length);
            Hash = hash;
        }

        /// <summary>
        /// Gets the file name, including extension.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the file full local path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the file full online path.
        /// </summary>
        public string Destination { get; }

        /// <summary>
        /// Gets the file length as long.
        /// </summary>
        public long Length { get; }

        /// <summary>
        /// Gets or sets the file upload progress.
        /// </summary>
        public double Progress { get; set; }

        /// <summary>
        /// Gets the file human-readable size.
        /// </summary>
        public string Size { get; }

        /// <summary>
        /// Indicates if the file size exceeds the OneDrive upload limit and therefore it needs zipping.
        /// </summary>
        public bool NeedsZip { get; }

        /// <summary>
        /// Gets or sets the Hash value
        /// </summary>
        public int Hash { get; set; }

        /// <summary>
        /// Check if file size exceeds the OneDrive upload limit for a single file.
        /// </summary>
        /// <param name="length">File length as long.</param>
        /// <returns>True if it exceeds the limit, otherwise false</returns>
        private static bool SizeExceedsLimit(long length)
        {
            long absolute = length < 0 ? -length : length;

            if (absolute >= 0x40000000) // Gigabyte
            {
                double readable = (length >> 20);
                readable = (readable / 1024);

                return readable >= 10;
            }

            return false;
        }
    }
}

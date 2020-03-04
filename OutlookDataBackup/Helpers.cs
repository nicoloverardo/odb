namespace OutlookDataBackup
{
    public static class Helpers
    {
        /// <summary>
        /// Get the human-readable file size for an arbitrary, 64-bit file size. 
        /// </summary>
        /// <param name="length">File length as long.</param>
        /// <returns>A System.String.</returns>
        /// <remarks>Code taken from <see cref="https://stackoverflow.com/a/11124118"/></remarks>
        public static string GetReadableSize(long length)
        {
            // Get absolute value
            long absolute_i = (length < 0 ? -length : length);

            // Determine the suffix and readable value
            string suffix;
            double readable;

            if (absolute_i >= 0x1000000000000000) // Exabyte
            {
                suffix = "EB";
                readable = (length >> 50);
            }
            else if (absolute_i >= 0x4000000000000) // Petabyte
            {
                suffix = "PB";
                readable = (length >> 40);
            }
            else if (absolute_i >= 0x10000000000) // Terabyte
            {
                suffix = "TB";
                readable = (length >> 30);
            }
            else if (absolute_i >= 0x40000000) // Gigabyte
            {
                suffix = "GB";
                readable = (length >> 20);
            }
            else if (absolute_i >= 0x100000) // Megabyte
            {
                suffix = "MB";
                readable = (length >> 10);
            }
            else if (absolute_i >= 0x400) // Kilobyte
            {
                suffix = "KB";
                readable = length;
            }
            else
            {
                return length.ToString("0 bytes"); // Byte
            }

            // Divide by 1024 to get fractional value
            readable = (readable / 1024);

            // Return formatted number with suffix
            return readable.ToString("0.## ") + suffix;
        }
    }
}

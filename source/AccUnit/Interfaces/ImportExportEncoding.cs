using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    /// <summary>
    /// Encoding of exported test class files (.cls, .bas).
    /// The import always detects the encoding (UTF-8 with or without BOM, otherwise ANSI).
    /// </summary>
    [ComVisible(true)]
    [Guid("FA0E0FD9-F0F3-40CB-9D8D-D7FD4502A8C0")]
    public enum ImportExportEncoding
    {
        Utf8Bom = 0,
        Ansi = 1
    }
}

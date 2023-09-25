// ReSharper disable IdentifierTypo
// ReSharper disable InconsistentNaming
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedType.Global
// ReSharper disable StringLiteralTypo
// ReSharper disable NotAccessedField.Global
// ReSharper disable RedundantAssignment
// ReSharper disable UnusedMember.Global

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

#nullable disable

using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace HamedStack.WindowsRawPrint;

/// <summary>
/// A utility class for sending data to a Windows printer using raw printing.
/// </summary>
[SuppressMessage("Interoperability", "CA1401:P/Invokes should not be visible")]
[SuppressMessage("Interoperability",
    "SYSLIB1054:Use \'LibraryImportAttribute\' instead of \'DllImportAttribute\' to generate P/Invoke marshalling code at compile time")]
public static class WindowsRawPrintUtility
{
    /// <summary>
    /// Represents document information for printing.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class DOCINFOA
    {
        /// <summary>
        /// Pointer to the data type name.
        /// </summary>
        [MarshalAs(UnmanagedType.LPStr)] public string pDataType;

        /// <summary>
        /// Pointer to the document name.
        /// </summary>
        [MarshalAs(UnmanagedType.LPStr)] public string pDocName;

        /// <summary>
        /// Pointer to the output file name.
        /// </summary>
        [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
    }

    /// <summary>
    /// Closes a printer connection.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer to be closed.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    /// <summary>
    /// Ends a print job.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);


    /// <summary>
    /// Ends a page within a print job.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    /// <summary>
    /// Opens a printer for printing.
    /// </summary>
    /// <param name="szPrinter">The name of the printer to be opened.</param>
    /// <param name="hPrinter">[out] A pointer to a handle to receive the printer handle if the function succeeds.</param>
    /// <param name="pd">Reserved. Must be IntPtr.Zero.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter,
        IntPtr pd);


    /// <summary>
    /// Sends a byte array to the specified printer.
    /// </summary>
    /// <param name="printerName">The name of the printer.</param>
    /// <param name="pBytes">A pointer to the byte array to be printed.</param>
    /// <param name="dwCount">The number of bytes to be printed.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    public static bool SendBytesTo(string printerName, IntPtr pBytes, int dwCount)
    {
        if (!IsWindowsOperatingSystem())
        {
            throw new InvalidOperationException("This library is designed to run on Windows operating systems only.");
        }

        // ReSharper disable once NotAccessedVariable
        var dwError = 0;
        // ReSharper disable once InlineOutVariableDeclaration
        IntPtr hPrinter;
        var di = new DOCINFOA();
        var bSuccess = false; // Assume failure unless you specifically succeed.

        di.pDocName = "My C#.NET RAW Document";
        di.pDataType = "RAW";

        // Open the printer.
        if (OpenPrinter(printerName.Normalize(), out hPrinter, IntPtr.Zero))
        {
            // Start a document.
            if (StartDocPrinter(hPrinter, 1, di))
            {
                // Start a page.
                if (StartPagePrinter(hPrinter))
                {
                    // Write your bytes.
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out _);
                    EndPagePrinter(hPrinter);
                }

                EndDocPrinter(hPrinter);
            }

            ClosePrinter(hPrinter);
        }

        // If you did not succeed, GetLastError may give more information about why not.
        if (bSuccess == false)
        {
            dwError = Marshal.GetLastWin32Error();
        }

        return bSuccess;
    }

    /// <summary>
    /// Sends a file to the specified printer.
    /// </summary>
    /// <param name="printerName">The name of the printer.</param>
    /// <param name="fileName">The path to the file to be printed.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    public static bool SendFileTo(string printerName, string fileName)
    {
        if (!IsWindowsOperatingSystem())
        {
            throw new InvalidOperationException("This library is designed to run on Windows operating systems only.");
        }

        // Open the file.
        var fs = new FileStream(fileName, FileMode.Open);
        return SendStreamTo(printerName, fs);
    }

    /// <summary>
    /// Sends the contents of a <see cref="Stream"/> to the specified printer.
    /// </summary>
    /// <param name="printerName">The name of the printer to which the stream should be sent.</param>
    /// <param name="stream">The <see cref="Stream"/> containing the data to be printed.</param>
    /// <returns><c>true</c> if the operation is successful, <c>false</c> otherwise.</returns>
    public static bool SendStreamTo(string printerName, Stream stream)
    {
        if (!IsWindowsOperatingSystem())
        {
            throw new InvalidOperationException("This library is designed to run on Windows operating systems only.");
        }

        // Create a BinaryReader on the stream.
        var br = new BinaryReader(stream);
        // Allocate an array of bytes to hold the stream's contents.
        var bytes = new byte[stream.Length];
        var bSuccess = false;
        // Initialize an unmanaged pointer.
        var pUnmanagedBytes = new IntPtr(0);

        var nLength = Convert.ToInt32(stream.Length);
        // Read the contents of the stream into the byte array.
        bytes = br.ReadBytes(nLength);
        // Allocate unmanaged memory for the byte array.
        pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
        // Copy the managed byte array into the unmanaged memory.
        Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
        // Send the unmanaged bytes to the printer.
        bSuccess = SendBytesTo(printerName, pUnmanagedBytes, nLength);
        // Free the unmanaged memory allocated earlier.
        Marshal.FreeCoTaskMem(pUnmanagedBytes);
        return bSuccess;
    }

    /// <summary>
    /// Sends a string to the specified printer.
    /// </summary>
    /// <param name="printerName">The name of the printer.</param>
    /// <param name="str">The string to be printed.</param>
    /// <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
    public static bool SendStringTo(string printerName, string str)
    {
        if (!IsWindowsOperatingSystem())
        {
            throw new InvalidOperationException("This library is designed to run on Windows operating systems only.");
        }

        // How many characters are in the string?
        var dwCount = str.Length;
        // Assume that the printer is expecting ANSI text, and then convert
        // the string to ANSI text.
        var pBytes = Marshal.StringToCoTaskMemAnsi(str);
        // Send the converted ANSI string to the printer.
        SendBytesTo(printerName, pBytes, dwCount);
        Marshal.FreeCoTaskMem(pBytes);
        return true;
    }

    /// <summary>
    /// Starts a new print job on the specified printer.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer.</param>
    /// <param name="level">The level of the structure being used.</param>
    /// <param name="di">A pointer to a <see cref="DOCINFOA"/> structure that describes the document's properties.</param>
    /// <returns><c>true</c> if the operation is successful, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, int level,
        [In, MarshalAs(UnmanagedType.LPStruct)]
        DOCINFOA di);

    /// <summary>
    /// Starts a new page within the current print job.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer.</param>
    /// <returns><c>true</c> if the operation is successful, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    /// <summary>
    /// Writes data to the specified printer.
    /// </summary>
    /// <param name="hPrinter">A handle to the printer.</param>
    /// <param name="pBytes">A pointer to the byte array to be written.</param>
    /// <param name="dwCount">The number of bytes to be written.</param>
    /// <param name="dwWritten">[out] The number of bytes actually written to the printer.</param>
    /// <returns><c>true</c> if the operation is successful, <c>false</c> otherwise.</returns>
    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

    /// <summary>
    /// Checks if the current operating system is Windows.
    /// </summary>
    /// <returns><c>true</c> if the current OS is Windows, <c>false</c> otherwise.</returns>
    private static bool IsWindowsOperatingSystem()
    {
        return Environment.OSVersion.Platform == PlatformID.Win32NT;
    }
}
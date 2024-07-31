using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine(">>    Press Enter to begin the installation   <<");
        Console.WriteLine("------------------------------------------------");
        Console.ReadLine();
        Console.Clear();

        string installPath = GetValidInstallationPath();
        string downloadUrl = "https://archive.org/download/cod-2-windows-mobile/COD2%20Windows%20Mobile.zip";
        string tempFilePath = Path.Combine(Path.GetTempPath(), "COD2 Windows Mobile.zip");

        if (!Directory.Exists(installPath + "\\COD2WME")) {
            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(">>          Downloading COD2: WME             <<");
            Console.WriteLine("------------------------------------------------");
            Console.Beep();

            await DownloadFileWithProgressAsync(downloadUrl, tempFilePath);

            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(">>          Extracting COD2: WME              <<");
            Console.WriteLine("------------------------------------------------");
            Console.Beep();

            // Extract the file to the installation path
            ZipFile.ExtractToDirectory(tempFilePath, installPath);

            // Delete the temporary file
            File.Delete(tempFilePath);
        }

        

        string realInstallPath = Path.Combine(installPath, "COD2WME", "MDEI");

        Console.Clear();
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("      Installing Microsoft Device Emulator      ");
        Console.WriteLine("------------------------------------------------");
        Console.Beep();

        // Determine system architecture
        
        bool is64Bit = Environment.Is64BitOperatingSystem;
        string exeFileName = is64Bit ? "vs_emulator_x64_vista.exe" : "vs_emulator.exe";
        string installerPath = Path.Combine(realInstallPath,  exeFileName);

        // Check if installer exists and run it
        if (File.Exists(installerPath))
        {
            Console.Clear();
            if (installerPath.Contains("x64"))
            {
                
                Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
                Console.WriteLine("------------------------------------------------");
                Console.WriteLine(" Installing Microsoft Device Emulator  (x64Bit) ");
                Console.WriteLine("------------------------------------------------");
            }
            else
            {
                
                Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
                Console.WriteLine("------------------------------------------------");
                Console.WriteLine(" Installing Microsoft Device Emulator  (x32Bit) ");
                Console.WriteLine("------------------------------------------------");
            }
            
                // Use PowerShell to run the executable with administrative privileges
                string powershellCommand = $"Start-Process -FilePath \"{installerPath}\" -Verb RunAs";
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-Command \"{powershellCommand}\"",
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    UseShellExecute = false,
                    Verb = "runAs",
                    CreateNoWindow = false
                };

                using (Process process = Process.Start(startInfo))
                {
                    using (StreamReader reader = process.StandardOutput)
                    {
                        string result = reader.ReadToEnd();
                        Console.WriteLine(result);
                    }
                }
            
            

        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: Installer not found at {installerPath}");
            
        }
        Console.Clear();
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("       Installed Microsoft Device Emulator      ");
        Console.WriteLine("------------------------------------------------");
        Console.Beep();
        Console.ReadLine();
    }

    static string GetValidInstallationPath()
    {
        string path;
        int attempts = 0;
        do
        {
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(">>        Type the path to install to         <<");
            Console.WriteLine("------------------------------------------------");
            Console.Write(">> ");
            path = Console.ReadLine().Trim(); // Read and trim whitespace

            if (!Directory.Exists(path))
            {
                attempts++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Clear();
                Console.WriteLine($"Error: The specified path does not exist. [ x{attempts} ]");

            }
            else
            {
                string confirmation = "";
                while (confirmation != "yes" && confirmation != "no")
                {
                    Console.Clear();
                    Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"Path Selected: {path}\\COD2 Windows Mobile ");
                    Console.WriteLine("Is this the correct path? (Yes/No)");
                    Console.WriteLine("------------------------------------------------");
                    Console.Write(">> ");
                    confirmation = Console.ReadLine().Trim().ToLower();
                }

                if (confirmation != "yes")
                {
                    attempts++;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.Clear();
                    Console.WriteLine("Please enter the installation path again.");

                    path = null; // Force the loop to continue
                }
            }
        } while (string.IsNullOrEmpty(path) || !Directory.Exists(path));

        return path;
    }

    static async Task DownloadFileWithProgressAsync(string url, string outputPath)
    {
        using (var client = new HttpClient())
        using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
        {
            response.EnsureSuccessStatusCode();
            long? totalBytes = response.Content.Headers.ContentLength;
            using (var contentStream = await response.Content.ReadAsStreamAsync())
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
            {
                var buffer = new byte[8192];
                int bytesRead;
                long totalBytesRead = 0;
                DateTime startTime = DateTime.Now;

                while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await fileStream.WriteAsync(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;

                    if (totalBytes.HasValue)
                    {
                        double progress = (double)totalBytesRead / totalBytes.Value;
                        double percentage = progress * 100;
                        double downloadSpeed = totalBytesRead / 1024.0 / 1024.0 / (DateTime.Now - startTime).TotalSeconds; // MB/s
                        bool kbTrue = false;
                        if (downloadSpeed < 1)
                        {
                            downloadSpeed = totalBytesRead / 1024.0 / (DateTime.Now - startTime).TotalSeconds; // KB/s
                            kbTrue = true;
                        }
                        double downloadedMB = totalBytesRead / 1024.0 / 1024.0;
                        double totalMB = totalBytes.Value / 1024.0 / 1024.0;

                        Console.Clear();
                        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition Installer");
                        Console.WriteLine("------------------------------------------------");
                        Console.WriteLine(">>          Downloading COD2: WME             <<");
                        Console.WriteLine($"Downloading... {percentage:0.00}%");
                        Console.WriteLine($"Downloaded: {downloadedMB:0.00} MB / {totalMB:0.00} MB");
                        if (kbTrue)
                        {
                            Console.WriteLine($"Speed: {downloadSpeed:0.00} KB/s");
                        }
                        else
                        {
                            Console.WriteLine($"Speed: {downloadSpeed:0.00} MB/s");
                        }
                        
                        Console.WriteLine("------------------------------------------------");
                    }
                }
            }
        }
    }
}

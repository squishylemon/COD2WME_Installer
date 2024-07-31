using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using IWshRuntimeLibrary;

class Program
{
    [STAThread]
    static async Task Main()
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine(">>    Press Enter to begin the installation   <<");
        Console.WriteLine("------------------------------------------------");
        Console.ReadLine();
        Console.Clear();

        string installPath = GetValidInstallationPath();
        string downloadUrl = "https://github.com/squishylemon/COD2WME_Installer/releases/download/1/CWMEI_SOURCE.zip";
        string tempFilePath = Path.Combine(Path.GetTempPath(), "CWMEI_SOURCE.zip");

        if (!Directory.Exists(installPath + "\\COD2WME")) {
            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(">>          Downloading COD2: WME             <<");
            Console.WriteLine("------------------------------------------------");
            Console.Beep();

            await DownloadFileWithProgressAsync(downloadUrl, tempFilePath);

            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(">>          Extracting COD2: WME              <<");
            Console.WriteLine("------------------------------------------------");
            Console.Beep();

            // Extract the file to the installation path
            ZipFile.ExtractToDirectory(tempFilePath, installPath);

            // Delete the temporary file
            System.IO.File.Delete(tempFilePath);
        }

        

        string realInstallPath = Path.Combine(installPath, "COD2WME");

        Console.Clear();
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("      Installing Microsoft Device Emulator      ");
        Console.WriteLine("------------------------------------------------");
        Console.Beep();

        bool is64Bit = Environment.Is64BitOperatingSystem;
        string bitFolder = is64Bit ? "64" : "32";
        string mdeiPath = Path.Combine(realInstallPath, "MDEI", bitFolder);
        string ppcFolderPath = Path.Combine(realInstallPath, "PPC");
        string deviceEmulatorExe = Path.Combine(mdeiPath, "DeviceEmulator.exe");
        string ppcUsaBin = Path.Combine(ppcFolderPath, "PPC_USA.bin");
        string pocketPcXml = Path.Combine(ppcFolderPath, "Pocket_pc", "Pocket_PC.xml");
        string shortcutPath = Path.Combine(realInstallPath, "COD2WME.lnk");
        // Ensure all required files exist
        if (System.IO.File.Exists(deviceEmulatorExe) && System.IO.File.Exists(ppcUsaBin) && System.IO.File.Exists(pocketPcXml))
        {
            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(is64Bit ? " Creating shortcut for Device Emulator (x64Bit) " : " Creating shortcut for Device Emulator (x32Bit) ");
            Console.WriteLine("------------------------------------------------");

            try
            {
                

                // Create a WshShell object to manage shortcuts
                WshShell wshShell = new WshShell();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(shortcutPath);

                shortcut.TargetPath = deviceEmulatorExe;
                shortcut.Arguments = $"\"{ppcUsaBin}\" /defaultsave /memsize 256 /sharedfolder \"{realInstallPath}\" /skin \"{pocketPcXml}\"";
                shortcut.WorkingDirectory = Path.GetDirectoryName(deviceEmulatorExe);
                shortcut.Description = "Shortcut to Cod2 via device emulator";
                shortcut.IconLocation = deviceEmulatorExe; // Use the executable's icon

                // Save the shortcut
                shortcut.Save();

                Console.Clear();
                Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
                Console.WriteLine("------------------------------------------------");
                Console.WriteLine("       Installed Microsoft Device Emulator      ");
                Console.WriteLine("------------------------------------------------");
                Console.Beep();

                Console.WriteLine($"Shortcut created COD2WME.");
                
            }
            catch (Exception ex)
            {
                Console.Clear();
                Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong  || " + "Error#" + ex.GetHashCode());
                Console.WriteLine("------------------------------------------------");
                Console.WriteLine("Failed to create shortcut. Make sure you have the necessary permissions.");
                Console.WriteLine("------------------------------------------------");
                Console.Beep();
                System.Threading.Thread.Sleep(100);
                Console.Beep();
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: One or more required files are missing.");
            if (!System.IO.File.Exists(deviceEmulatorExe))
            {
                Console.WriteLine($"deviceEmulatorExe - {deviceEmulatorExe}");
            }
            if (!System.IO.File.Exists(ppcUsaBin))
            {
                Console.WriteLine($"ppcUsaBin - {ppcUsaBin}");
            }
            if (!System.IO.File.Exists(pocketPcXml))
            {
                Console.WriteLine($"pocketPcXml - {pocketPcXml}");
            }
        }

        Console.Clear();
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("               Installing Save File             ");
        Console.WriteLine("------------------------------------------------");
        Console.Beep();
        string sourceFilePath = Path.Combine(realInstallPath, "SVE", "{00000000-0000-0000-0000-000000000000}.dess");
        string destinationDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Device Emulator");
        string destinationFilePath = Path.Combine(destinationDirectory, Path.GetFileName(sourceFilePath));

        try
        {
            // Check if the destination directory exists, create it if it doesn't
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
                
            }

            // Copy the file to the destination directory
            if (System.IO.File.Exists(sourceFilePath))
            {
                System.IO.File.Copy(sourceFilePath, destinationFilePath, true); // true to overwrite if the file already exists
                Console.Clear();
                Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
                Console.WriteLine("------------------------------------------------");
                Console.WriteLine("               Installed Save File              ");
                Console.WriteLine("------------------------------------------------");
                Console.Beep();
            }
            else
            {
                Console.WriteLine($"Source file does not exist: {sourceFilePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        Console.Clear();
        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("        Want to Run Cod: WME  (Yes or No)       ");
        Console.WriteLine("------------------------------------------------");
        Console.Write(">> ");
        Console.Beep();
        string confirmation2 = Console.ReadLine().Trim().ToLower();
        while (confirmation2 != "yes" && confirmation2 != "no")
        {
            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("        Want to Run Cod: WME  (Yes or No)       ");
            Console.WriteLine("------------------------------------------------");
            Console.Write(">> ");
            confirmation2 = Console.ReadLine().Trim().ToLower();
        }
        if (confirmation2 != "yes")
        {
            
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Clear();
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine(" Call Of Duty 2: Windows Mobile Edition Insalled");
            Console.WriteLine("------------------------------------------------");
            PlayInstallMusic();

        }else
        {
            try
            {
                // Start the shortcut
                Process.Start(new ProcessStartInfo
                {
                    FileName = shortcutPath,
                    UseShellExecute = true // UseShellExecute is needed for .lnk files
                });
                Console.WriteLine("Shortcut executed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Unable to execute shortcut. {ex.Message}");
            }
        }
    }

    static string GetValidInstallationPath()
    {
        string path;
        int attempts = 0;
        do
        {
            Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
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
                    Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"Path Selected: {path}\\COD2WME");
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
                        Console.WriteLine("Call Of Duty 2: Windows Mobile Edition\nMade For Jamievlong Check Him out at https://twitch.tv/jamievlong ");
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

    static void PlayInstallMusic()
    {
      
        int[] notes = { 440, 493, 523, 493, 440, 440, 440, 493, 493, 493, 440, 587, 587, 440 };
        int[] durations = { 200, 200, 200, 200, 200, 200, 200, 200, 200, 200, 400, 200, 200, 400 };

        for (int i = 0; i < notes.Length; i++)
        {
            Console.Beep(notes[i], durations[i]);
            Thread.Sleep(0); // Short pause between notes
        }
    }
}

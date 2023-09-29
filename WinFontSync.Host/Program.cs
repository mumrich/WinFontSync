using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;

string fontDirectory                            = "C:\\Users\\HMO\\OneDrive - INGTES AG\\Desktop\\Fonts";
InstalledFontCollection installedFontCollection = new();
PrivateFontCollection privateFontCollection     = new();
int nbrOfInstalledFonts                         = privateFontCollection.Families.Length;
IEnumerable<string> fontFilePaths               = Directory
                                                    .GetFiles(fontDirectory)
                                                    .Where(fp => fp.EndsWith(".ttf") || fp.EndsWith(".otf"));

foreach (string fontFilePath in fontFilePaths)
{
  privateFontCollection.AddFontFile(fontFilePath);

  if (privateFontCollection.Families.Length > nbrOfInstalledFonts)
  {
    nbrOfInstalledFonts = privateFontCollection.Families.Length;

    byte[] fontData = File.ReadAllBytes(fontFilePath);
    IntPtr fontPtr  = Marshal.AllocCoTaskMem(fontData.Length);

    // Installing the font
    Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
    privateFontCollection.AddMemoryFont(fontPtr, fontData.Length);
    Marshal.FreeCoTaskMem(fontPtr);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Font installed successfully: '{fontFilePath}'");
    Console.ForegroundColor = ConsoleColor.White;
  }
  else
  {
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Failed to install the font file: '{fontFilePath}'");
    Console.ForegroundColor = ConsoleColor.White;
  }
}

foreach (FontFamily font in installedFontCollection.Families)
{
  Console.WriteLine(font.Name);
}
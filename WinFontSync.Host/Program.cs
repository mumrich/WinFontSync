using System.Drawing;
using System.Drawing.Text;

InstalledFontCollection installedFontCollection = new();

foreach (FontFamily font in installedFontCollection.Families)
{
  Console.WriteLine(font.Name);
}
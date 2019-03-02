using CommandLine;

namespace CodeSanook.CollaborativeLocalization
{
    [Verb("export", HelpText = "Export a Google Spreadsheet localization to JSON file")]
    public class ExportOptions
    {
        [Option("output-dir", Required = true)]
        public string OutputDir { get; set; }

        [Option("key-to-upper-case", Required = false)]
        public bool UpdateKeyToUpperCase { get; set; }
    }
}

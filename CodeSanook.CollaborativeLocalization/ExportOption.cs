using CommandLine;
using System.Collections.Generic;

namespace CodeSanook.CollaborativeLocalization
{
    [Verb("export", HelpText = "Export a Google Spreadsheet localization to JSON file")]
    public class ExportOptions
    {
        [Option("output-dir", Required = true, HelpText = "A target directory for an output JSON file")]
        public string OutputDir { get; set; }

        [Option(
            "key-to-upper-case",
            Required = false,
            HelpText = "Should set a localization key to upper case automatically"
        )]
        public bool UpdateKeyToUpperCase { get; set; }

        // https://github.com/commandlineparser/commandline/wiki/CommandLine-Grammar#sequence-option
        // Sequence option
        [Option(
            "supported-languages",
            Required = false,
            HelpText = "Set a supported language for localization",
            Default = new[] { "en", "th" }
        )]
        public IEnumerable<string> SupportedLanguages { get; set; }

        [Option(
            "shared-to-emails",
            Required = false,
            HelpText = "Set emails to share a write permission to this sheet",
            Default = new string[] { }
        )]
        public IEnumerable<string> SharedToEmails { get; set; }

        [Option(
            "application-name",
            Required = false,
            HelpText = "Set Google API application name",
            Default = "CodeSanook.CollaborativeLocalization"
        )]
        public string ApplicationName { get; set; }

        [Option(
            "application-name",
            Required = false,
            HelpText = "Set service account email",
            Default = "collaborative-localization@codesanook.iam.gserviceaccount.com"
        )]
        public string ServiceAccountEmail { get; set; }

        [Option(
            "sheet-name",
            Required = false,
            HelpText = "Set Gooogle sheet name that you can find in Google drive",
            Default = "CodeSanook.CollaborativeLocalization"
        )]
        public string SheetName { get; set; }
    }
}


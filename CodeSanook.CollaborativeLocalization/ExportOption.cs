using CommandLine;
using System.Collections.Generic;

namespace CodeSanook.CollaborativeLocalization
{
    [Verb("export", HelpText = "Export a Google Spreadsheet localization sheet to JSON file")]
    public class ExportOptions
    {
        [Option(
            "output-dir",
            HelpText = "An ouput directory of JSON local file, default to a current working directory",
            Required = false,
            Default = "."
        )]
        public string OutputDir { get; set; }

        [Option(
            "sheet-name",
            Required = false,
            HelpText = "Gooogle sheet name, it will be created automatically if it does exist",
            Default = "Codesanook.CollaborativeLocalization"
        )]
        public string SheetName { get; set; }

        [Option(
            "key-to-upper-case",
            Required = false,
            HelpText = "Set a localization key to upper case automatically",
            Default = true
        )]
        public bool UpdateKeyToUpperCase { get; set; }

        // Sequence option https://github.com/commandlineparser/commandline/wiki/CommandLine-Grammar#sequence-option
        [Option(
            "supported-languages",
            Required = false,
            HelpText = "Space seperated value of supported language, these values match all sheet tabs",
            Default = new[] { "en", "th" }
        )]
        public IEnumerable<string> SupportedLanguages { get; set; }

        [Option(
            "shared-to-emails",
            Required = false,
            HelpText = "Space separated values emails to share a write permission to the working sheet",
            Default = new string[] { }
        )]
        public IEnumerable<string> SharedToEmails { get; set; }

        [Option(
            "service-account-email",
            Required = false,
            HelpText = "Service account email",
            Default = "collaborative-localization@codesanook.iam.gserviceaccount.com"
        )]
        public string ServiceAccountEmail { get; set; }

        [Option(
            "application-name",
            Required = false,
            HelpText = "Google API application name",
            Default = "Codesanook.CollaborativeLocalization"
        )]
        public string ApplicationName { get; set; }
    }
}


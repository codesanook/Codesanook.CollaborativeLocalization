using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using System.Linq;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using System;
using CommandLine;

namespace CodeSanook.CollaborativeLocalization
{
    public class Program
    {
        // Downloaded from https://console.developers.google.com  
        // Put to root rectory and set copy to output directory "copy if newer"
        const string keyFilePath = @"service-account-private-key.p12";

        private static DriveService driverService;
        private static SheetsService sheetService;

        public static void Main(string[] args)
        {
            Parser.Default.ParseArguments<ExportOptions>(args)
            .WithParsed(exportOptions => RunExportOptions(exportOptions))
            .WithNotParsed(errs => { Console.WriteLine(errs); });
        }

        private static void RunExportOptions(ExportOptions options)
        {
            // Load the Key file
            var certificate = new X509Certificate2(
                keyFilePath,
                "notasecret",
                X509KeyStorageFlags.Exportable
            );

            // If modifying these scopes, delete your previously saved credentials
            var credential = new ServiceAccountCredential(
                new ServiceAccountCredential.Initializer(options.ServiceAccountEmail)
                {
                    Scopes = new[] { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets },
                }.FromCertificate(certificate)
            );

            // Create Google driver service.
            driverService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = options.ApplicationName
            });

            // Create Google Sheets API service.
            sheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = options.ApplicationName,
            });

            var spreadsheetId = GetSpreadsheetId(options);
            // Create sheet if does not exist
            if (string.IsNullOrEmpty(spreadsheetId))
            {
                spreadsheetId = CreateSpreadsheet(options);
                InitialCellValues(spreadsheetId, options);
            }

            foreach (var sheetName in options.SupportedLanguages)
            {
                ExportToJson(spreadsheetId, sheetName, options);
            }
        }

        private static void ExportToJson(string spreadsheetId, string sheetName, ExportOptions options)
        {
            var request = sheetService.Spreadsheets.Values.Get(spreadsheetId, $"'{sheetName}'!A:B");
            var response = request.Execute();
            var values = response.Values;

            //Skip the first row as a header
            var localeData = new Dictionary<object, object>();
            for (var rowIndex = 1; rowIndex < values.Count; rowIndex++)
            {
                var row = values[rowIndex];
                if (row.Count != 2 || string.IsNullOrWhiteSpace(row[0]?.ToString()) || string.IsNullOrWhiteSpace(row[1]?.ToString()))
                {
                    throw new InvalidOperationException($"sheet name '{sheetName}', row number {rowIndex + 1} has empty key or value, please fix it and run the tool again.");
                }

                var key = row[0].ToString().Trim();
                var value = row[1].ToString().Trim();
                key = UpdateLocalizationKeyToUpperCaseIfRequired(spreadsheetId, sheetName, options, rowIndex, key);
                localeData.Add(key, value);
            }

            var exportedLocale = JsonConvert.SerializeObject(localeData, Formatting.Indented);
            var exportedFilePath = Path.GetFullPath(Path.Combine(options.OutputDir, $"{sheetName}.json"));
            var fileInfo = new FileInfo(exportedFilePath);
            fileInfo.Directory.Create();

            using (var fileStream = new FileStream(exportedFilePath, FileMode.Create, FileAccess.Write))
            using (var streamWriter = new StreamWriter(fileStream))
            {
                streamWriter.Write(exportedLocale);
            }
            Console.WriteLine($"{exportedFilePath} exported");
        }

        private static string UpdateLocalizationKeyToUpperCaseIfRequired(
            string spreadsheetId,
            string sheetName,
            ExportOptions exportOptions,
            int rowIndex,
            string localizationKey
        )
        {
            if (!exportOptions.UpdateKeyToUpperCase) return localizationKey;

            // Prevent unnecessary update and web request
            var anyLowerCaseCharacterInKey =
                localizationKey.ToArray()
                .Any(c => char.IsLetter(c) && char.IsLower(c));
            if (anyLowerCaseCharacterInKey)
            {
                localizationKey = localizationKey.ToUpper();
                UpdateLocalizationKey(spreadsheetId, sheetName, rowIndex, localizationKey);
            }
            return localizationKey;
        }

        private static void UpdateLocalizationKey(
            string spreadsheetId,
            string sheetName,
            int rowIndex,
            string localizationKey
        )
        {
            // Array which its member is a array (2 dimensions array)
            var valueRange = new ValueRange { Values = new string[][] { new[] { localizationKey } } };
            var cellRange = $"'{sheetName}'!A{rowIndex + 1}";

            var request = sheetService.Spreadsheets.Values.Update(valueRange, spreadsheetId, cellRange);
            request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }

        private static void InitialCellValues(string spreadsheetId, ExportOptions options)
        {
            var sheetNames = options.SupportedLanguages.ToArray();
            foreach (var sheetName in sheetNames)
            {
                UpdateHeader(spreadsheetId, sheetName, "A1:B1", new[] { "Localization Key", "Localization Value" });
            }

            // Auto create 500 cells on the first sheet 
            var rowCount = 500;
            // Take a localization keys from a first sheet tab and fill to other tabs
            var values = Enumerable.Range(2, rowCount)
                 .Select(value => new[] { $"='{sheetNames[0]}'!A{value}" })
                 .ToArray();
            var body = new ValueRange { Values = values };

            for (var sheetIndex = 1; sheetIndex < sheetNames.Length; sheetIndex++)
            {
                var request = sheetService.Spreadsheets.Values.Update(
                    body,
                    spreadsheetId,
                    $"'{sheetNames[sheetIndex]}'!A2:A{rowCount + 1}"
                );
                request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
                request.Execute();
            }
        }

        private static void UpdateHeader(string spreadsheetId, string sheetName, string cellRange, object[] headerTitles)
        {
            var body = new ValueRange { Values = new[] { headerTitles } };
            var request = sheetService.Spreadsheets.Values.Update(body, spreadsheetId, $"'{sheetName}'!{cellRange}");
            request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }

        private static string CreateSpreadsheet(ExportOptions options)
        {
            var sheets = options.SupportedLanguages.Select(title =>
               new Sheet() { Properties = new SheetProperties() { Title = title } }
            ).ToList();

            var spreadSheet = new Spreadsheet
            {
                Properties = new SpreadsheetProperties() { Title = options.SheetName },
                Sheets = sheets
            };
            var createSpreadSheetResponse = sheetService.Spreadsheets.Create(spreadSheet).Execute();

            ShareWritePerssion(createSpreadSheetResponse, options);
            return createSpreadSheetResponse.SpreadsheetId;
        }

        private static void ShareWritePerssion(Spreadsheet createSpreadSheetResponse, ExportOptions options)
        {
            foreach (var email in options.SharedToEmails)
            {
                // Update permission
                var file = driverService.Files.Get(createSpreadSheetResponse.SpreadsheetId).Execute();
                var permission = new Permission
                {
                    Role = "writer",
                    Type = "user",
                    EmailAddress = email
                };

                //request.TransferOwnership = true; // Work only same domain email, e.g google app for business
                var request = driverService.Permissions.Create(permission, file.Id);
                request.Execute();
            }
        }

        private static string GetSpreadsheetId(ExportOptions option)
        {
            var listRequest = driverService.Files.List();
            //https://developers.google.com/drive/api/v3/reference/query-ref
            listRequest.Q = $"name='{option.SheetName}' and mimeType='application/vnd.google-apps.spreadsheet'";
            var files = listRequest.Execute().Files;
            return files.SingleOrDefault(f => f.Name == option.SheetName)?.Id;
        }
    }
}

// Tips create and delete a file

//var file = new File();
//file.Name = applicationName;
//file.MimeType = "application/vnd.google-apps.spreadsheet";
//var insert = driverService.Files.Create(file);
//var response = file = insert.Execute();

//for (int index = 0; index < files.Count; index++)
//{
//    driverService.Files.Delete(files[index].Id).Execute();
//}

//var fileResponse =  driverService.Files.Get(response.Id).Execute();


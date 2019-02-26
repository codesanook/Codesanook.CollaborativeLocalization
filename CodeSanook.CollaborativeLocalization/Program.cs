using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using System.Linq;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest;

namespace CodeSanook.CollaborativeLocalization
{
    public class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        const string applicationName = "CodeSanook.CollaborativeLocalization";
        const string serviceAccountEmail = "collaborative-localization@codesanook.iam.gserviceaccount.com";

        // Downloaded from https://console.developers.google.com 
        const string keyFilePath = @"service-account-private-key.p12";
        private static DriveService driverService;
        private static SheetsService sheetService;
        private static string[] sheetNames = new[] { "en", "th" };
        private static string[] sharedEmails = new [] { "" };

        static void Main(string[] args)
        {
            //loading the Key file
            var certificate = new X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable);
            var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = new[] { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets },

            }.FromCertificate(certificate));

            driverService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });

            // Create Google Sheets API service.
            sheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });

            var spreadsheetId = GetSpreadsheetId();
            if (string.IsNullOrEmpty(spreadsheetId))
            {
                spreadsheetId = CreateSpreadsheet();
            }

            UpdateCellValues(spreadsheetId);
        }

        private static void UpdateCellValues(string spreadsheetId)
        {
            UpdateHeader(spreadsheetId, sheetNames[0], "A1:B1", new[] { "Key", "Value" });
            UpdateHeader(spreadsheetId, sheetNames[1], "A1:B1", new[] { "Key", "Value" });

            var rowCount = 500;
            var values = Enumerable.Range(2, rowCount).Select(value => new[] { $"='{sheetNames[0]}'!A{value}" }).ToArray();
            var body = new ValueRange { Values = values };

            var request = sheetService.Spreadsheets.Values.Update(body, spreadsheetId, $"'{sheetNames[1]}'!A2:A{rowCount + 1}");
            request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
        }

        private static void UpdateHeader(string spreadsheetId, string sheetName, string cellRange, object[] headerTitles)
        {
            var body = new ValueRange
            {
                Values = new[]
                {
                    headerTitles
                }
            };
            var request = sheetService.Spreadsheets.Values.Update(body, spreadsheetId, $"'{sheetName}'!{cellRange}");
            request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
        }

        private static string CreateSpreadsheet()
        {
            var sheets = sheetNames.Select(title =>
               new Sheet() { Properties = new SheetProperties() { Title = title } }
            ).ToList();

            var spreadSheet = new Spreadsheet
            {
                Properties = new SpreadsheetProperties() { Title = applicationName },
                Sheets = sheets
            };
            var createSpreadSheetResponse = sheetService.Spreadsheets.Create(spreadSheet).Execute();

            //update permission
            var file = driverService.Files.Get(createSpreadSheetResponse.SpreadsheetId).Execute();
            var permission = new Permission();
            permission.Role = "writer";
            permission.Type = "user";
            permission.EmailAddress = sharedEmails[0];
            var request = driverService.Permissions.Create(permission, file.Id);
            //request.TransferOwnership = true; //work only same domain
            request.Execute();
            return createSpreadSheetResponse.SpreadsheetId;
        }

        private static string GetSpreadsheetId()
        {
            var listRequest = driverService.Files.List();
            //https://developers.google.com/drive/api/v3/reference/query-ref
            listRequest.Q = $"name='{applicationName}' and mimeType='application/vnd.google-apps.spreadsheet'";
            var files = listRequest.Execute().Files;
            return files.SingleOrDefault(f => f.Name == applicationName)?.Id;
        }
    }
}

//for (int index = 0; index < files.Count; index++)
//{
//    driverService.Files.Delete(files[index].Id).Execute();
//}

//return;
//var file = new File();
//file.Name = applicationName;
//file.MimeType = "application/vnd.google-apps.spreadsheet";
//var insert = driverService.Files.Create(file);
//var response = file = insert.Execute();

//var fileResponse =  driverService.Files.Get(response.Id).Execute();
// fileResponse.



/*
*/


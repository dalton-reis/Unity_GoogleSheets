using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace TecEduFURB.GoogleSpreadsheet {

    public class SpreadsheetService {
        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private readonly IList<Sheet> sheets;

        private int sheetNumber;
        private string spreadsheetId;
        private string credentialPath;

        private SheetsService service;

        public SpreadsheetService(string credentialPath, string spreadsheetId, int sheetNumber = 0) {
            this.sheetNumber = sheetNumber;
            this.spreadsheetId = spreadsheetId;
            this.credentialPath = credentialPath;

            CreateService(credentialPath);
            sheets = GetSheets();

            if (sheetNumber > sheets.Count)
                sheetNumber = sheets.Count;
        }

        private void CreateService(string credentialPath) {
            UserCredential credential = CreateCredential(credentialPath);
            service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential
            });
        }

        private UserCredential CreateCredential(string credentialPath) {
            UserCredential credential;

            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read)) {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                ).Result;
            }

            return credential;
        }

        private IList<Sheet> GetSheets() {
            SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(spreadsheetId);

            Spreadsheet response = request.Execute();
            IList<Sheet> values = response.Sheets;

            return response.Sheets;
        }

        public IList<IList<object>> GetHeaders() {
            return GetValues("!1:1");
        }

        public IList<IList<object>> GetValues(string range = null) {
            if (range == null)
                range = sheets[sheetNumber].Properties.Title;

            GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();

            return response.Values;
        }

        public AppendValuesResponse AddValues(List<object> values, string range) {
            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { values };

            AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;

            return appendRequest.Execute();
        }

        public UpdateValuesResponse UpdateCell(string cell, object newValue) {
            List<object> oblist = new List<object>() { newValue };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { oblist };

            UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, cell);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;

            return updateRequest.Execute();
        }

        public ClearValuesResponse DeleteValues(string range = null) {
            if (range == null)
                range = sheets[sheetNumber].Properties.Title;

            ClearValuesRequest requestBody = new ClearValuesRequest();

            ClearRequest deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            return deleteRequest.Execute();
        }
    }
}

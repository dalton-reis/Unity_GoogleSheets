namespace TecEduFURB.GoogleSheets {

    /// <summary>
    /// Contém as informações referêntes às credenciais utilizadas pelo SpreadsheetService.
    /// Acesse o guia da Google Sheets API para mais informações: 
    /// https://developers.google.com/sheets/api/guides/concepts.
    /// 
    /// Autor: github.com/AlexSerodio
    /// </summary>
    public abstract class SpreadsheetInfo {

        public const string CREDENTIALS_PATH = "./Assets/GoogleSheetsAPi/credentials.json";
        public const string SPREADSHEET_ID = "colocar aqui id da planilha do Google disponível na url da planilha";

    }

}
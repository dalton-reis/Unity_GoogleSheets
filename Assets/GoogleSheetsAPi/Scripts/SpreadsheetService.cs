using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace TecEduFURB.GoogleSheets {

    /// <summary>
    /// Disponibiliza métodos para interagir com planilhas do Google.
    /// Acesse o guia da Google Sheets API para mais informações: 
    /// https://developers.google.com/sheets/api/guides/concepts.
    /// 
    /// Autor: github.com/AlexSerodio
    /// </summary>
    public class SpreadsheetService {
        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private readonly IList<Sheet> sheets;

        private int sheetNumber;
        private string spreadsheetId;
        private string credentialPath;

        private SheetsService service;

        /// <summary>
        /// Construtor padrão para criação do service.
        /// </summary>
        /// <param name="credentialPath">Caminho do arquivo com a credencial gerada pelo Google API Console.</param>
        /// <param name="spreadsheetId">ID da planilha disponível na url da mesma.</param>
        /// <param name="sheetNumber">Opcional - Index da folha/aba desejada da planilha (default = 0).</param>
        public SpreadsheetService(string credentialPath, string spreadsheetId, int sheetNumber = 0) {
            if(string.IsNullOrEmpty(credentialPath))
                throw new ArgumentException("O parâmetro credentialPath deve ser um valor não null e não vazio.");
            if(string.IsNullOrEmpty(spreadsheetId))
                throw new ArgumentException("O parâmetro spreadsheetId deve ser um valor não null e não vazio.");
            
            this.sheetNumber = sheetNumber;
            this.spreadsheetId = spreadsheetId;
            this.credentialPath = credentialPath;

            CreateService(credentialPath);
            sheets = GetSheets();

            if (sheetNumber > sheets.Count)
                sheetNumber = sheets.Count;
        }

        /// <summary>
        /// Cria uma nova instância do SheetsService com as credenciais definidas.
        /// </summary>
        /// <param name="credentialPath">Caminho do arquivo com a credencial gerada pelo Google API Console.</param>
        private void CreateService(string credentialPath) {
            UserCredential credential = CreateCredential(credentialPath);
            service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential
            });
        }

        /// <summary>
        /// Cria o arquivo 'token.json' contendo as credenciais de usuário necessárias para o serviço.
        /// </summary>
        /// <param name="credentialPath">Caminho do arquivo com a credencial gerada pelo Google API Console.</param>
        /// <returns>Objeto contendo as credenciais do usuário.</returns>
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

        /// <summary>
        /// Recupera todas as folhas/abas da planilha.
        /// </summary>
        /// <returns>Uma lista contendo todas as abas da planilha.</returns>
        private IList<Sheet> GetSheets() {
            SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(spreadsheetId);

            Spreadsheet response = request.Execute();
            IList<Sheet> values = response.Sheets;

            return response.Sheets;
        }

        /// <summary>
        /// Recupera apenas o cabeçalho da planilha (primeira linha).
        /// </summary>
        /// <returns>O cabeçalho da planilha.</returns>
        public IList<IList<object>> GetHeaders() {
            return GetValues("!1:1");
        }

        /// <summary>
        /// Recupera os itens da planilha que estejam dentro do range informado
        /// ou todos os itens caso nenhum range seja informado.
        /// </summary>
        /// <param name="range">Opcional - O range contendo os itens que deseja retornar. Exemplo: A1:B2.</param>
        /// <returns>Lista contendo todos os itens recuperados.</returns>
        public IList<IList<object>> GetValues(string range = null) {
            if (range == null)
                range = sheets[sheetNumber].Properties.Title;

            GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();

            return response.Values;
        }

        /// <summary>
        /// Adiciona os valores informados na planilha.
        /// </summary>
        /// <param name="values">Os valores da linha a ser adicionada.</param>
        /// <returns>Resposta do request da operação.</returns>
        public AppendValuesResponse AddValues(List<object> values) {
            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { values };

            AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, "A:A");
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;

            return appendRequest.Execute();
        }

        /// <summary>
        /// Atualiza o valor de uma célula da planilha.
        /// </summary>
        /// <param name="cell">Índice da célula a ser atualizada. Exemplo: A2.</param>
        /// <param name="newValue">Novo valor da célula.</param>
        /// <returns>Resposta do request da operação.</returns>
        public UpdateValuesResponse UpdateCell(string cell, object newValue) {
            List<object> oblist = new List<object>() { newValue };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { oblist };

            UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, cell);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;

            return updateRequest.Execute();
        }

        /// <summary>
        /// Remove um ou mais itens da planilha.
        /// </summary>
        /// <param name="range">Range dos itens a serem removidos. Exemplo: A1:B2.</param>
        /// <returns>Resposta do request da operação.</returns>
        public ClearValuesResponse DeleteValues(string range = null) {
            if (range == null)
                range = sheets[sheetNumber].Properties.Title;

            ClearValuesRequest requestBody = new ClearValuesRequest();

            ClearRequest deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            return deleteRequest.Execute();
        }
    }
}

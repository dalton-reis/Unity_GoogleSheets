using System.Collections.Generic;
using TecEduFURB.GoogleSheets;
using UnityEngine;

/// <summary>
/// Demonstra um caso de uso da classe SpreadsheetService, no qual são recuperado todos os valores
/// da planilha definida pela classe SpreadsheetInfo.
/// 
/// Autor: github.com/AlexSerodio
/// </summary>
public class TestSpreadsheetService : MonoBehaviour {

    private SpreadsheetService service;

    void Start() {
        service = new SpreadsheetService(SpreadsheetInfo.CREDENTIALS_PATH, SpreadsheetInfo.SPREADSHEET_ID);

        service.AddValues(new List<object>(){"elemento 1", "elemento 2"});

        service.UpdateCell("A2", "atualizado");

        IList<IList<object>> values = service.GetValues("A2:B2");

        foreach(var row in values) {
            foreach (var col in row)
                Debug.Log(col);
        }

    }
}

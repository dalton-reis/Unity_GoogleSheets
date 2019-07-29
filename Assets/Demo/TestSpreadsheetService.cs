﻿using System.Collections.Generic;
using UnityEngine;

namespace TecEduFURB.GoogleSpreadsheet {

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

            IList<IList<object>> values = service.GetValues();

            foreach(var row in values) {
                foreach (var col in row)
                    Debug.Log(col);
            }

        }
    }
}
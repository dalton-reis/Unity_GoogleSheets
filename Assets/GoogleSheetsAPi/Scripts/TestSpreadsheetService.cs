using System.Collections.Generic;
using UnityEngine;

namespace TecEduFURB.GoogleSpreadsheet {

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

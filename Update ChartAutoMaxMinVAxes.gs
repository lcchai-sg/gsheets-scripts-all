// This function passes a simulated event to the onEdit trigger function,
// which allows testing of trigger using GAS.
// function test_onEdit() {
//     onEdit({
//         user: Session.getActiveUser().getEmail(),
//         source: SpreadsheetApp.getActiveSpreadsheet(),
//         range: SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
//         value: SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
//         authMode: "LIMITED"
//     });
// }

// Trigger function to automatically adjust chart vertical axes for a set of sheets.
function onEdit(e) {
    var buffer = 0.25; // buffer for max/min values
    var max = null;
    var min = null;
    var sheet = e.source.getActiveSheet();
    var values = [];

    if (sheet.getName() == "Value By Category" || sheet.getName() == "Value By Item") {
        // max/min are max/min values from columns G and J
        values = sheet.getRange("G:G").getValues().filter(Number);
        values = values.concat(sheet.getRange("J:J").getValues().filter(Number));
    } else if (sheet.getName() == "Total Value") {
        // max/min are max/min values from columns B and D
        values = sheet.getRange("B:B").getValues().filter(Number);
        values = values.concat(sheet.getRange("D:D").getValues().filter(Number));
    }

    if (values.length) {
        max = Math.max(...values);
        max += Math.abs(max) * buffer;
        min = Math.min(...values);
        min -= Math.abs(min) * buffer;

        var chart = sheet.getCharts()[0];

        chart = chart.modify()
            .setOption('vAxes.0.viewWindow.max', max)
            .setOption('vAxes.0.viewWindow.min', min)
            .build();

        sheet.updateChart(chart);
    }
}

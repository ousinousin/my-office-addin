Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Excel is ready");
    }
});

document.getElementById("run").addEventListener("click", () => {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [["Hello, Office Add-ins!"]];
        await context.sync();
    });
});

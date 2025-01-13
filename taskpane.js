Office.onReady((info) => {
    // アドインが読み込まれたときに呼び出される
    if (info.host === Office.HostType.Word) {
        document.getElementById("run").onclick = () => tryInsertText();
    }
});

async function tryInsertText() {
    try {
        await Word.run(async (context) => {
            const docBody = context.document.body;
            docBody.insertText("Hello, Word Add-in!", Word.InsertLocation.end);
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Word API バージョン: ", Office.context.requirements.isSetSupported('WordApi', '1.3'));
    }
});
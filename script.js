document.getElementById("insertText").onclick = async () => {
    try {
        await Word.run(async (context) => {
            // ドキュメントの末尾にテキストを挿入
            const paragraph = context.document.body.insertParagraph(
                "Hello from My Office Add-in!",
                Word.InsertLocation.end
            );
            paragraph.font.color = "blue"; // テキストの色を青に変更
            await context.sync(); // 変更を反映
        });
    } catch (error) {
        console.error("エラーが発生しました:", error);
    }
};

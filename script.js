document.getElementById("insertText").onclick = async () => {
    try {
        await Word.run(async (context) => {
            const paragraph = context.document.body.insertParagraph(
                "Hello from My Add-in!",
                Word.InsertLocation.end
            );
            paragraph.font.color = "blue";
            await context.sync();
        });
    } catch (error) {
        console.error("エラー:", error);
    }
};

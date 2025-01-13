const mainJs = `Office.onReady(() => {
  document.getElementById('insertText').addEventListener('click', () => {
    Word.run(async (context) => {
      const doc = context.document;
      const selection = doc.getSelection();
      selection.insertText("Hello, Word Online!", Word.InsertLocation.replace);
      await context.sync();
    });
  });
});`;

export { manifestXml, indexHtml, mainJs };
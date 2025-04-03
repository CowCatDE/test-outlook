(function() {
    'use strict';

    Office.onReady(function() {
        // Office ist bereit
        document.getElementById('insertHtmlButton').onclick = insertHtmlAsText;
    });

    function escapeHtml(html) {
        // HTML escapen, damit es als Text und nicht als HTML interpretiert wird
        return html
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function insertHtmlAsText() {
        var htmlCode = document.getElementById('htmlCode').value;
        var escapedHtml = escapeHtml(htmlCode);
        
        // Fügt den escaped HTML-Code in die E-Mail ein
        // Wir setzen die Schriftart auf Courier New (oder eine andere Monospace-Schrift),
        // um den Code besser lesbar zu machen
        var htmlToInsert = '<pre style="font-family: Courier New, monospace;">' + escapedHtml + '</pre>';
        
        // In die E-Mail einfügen
        Office.context.mailbox.item.body.setSelectedDataAsync(
            htmlToInsert,
            { coercionType: Office.CoercionType.Html },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Fehler beim Einfügen des HTML-Codes:', result.error.message);
                } else {
                    console.log('HTML-Code erfolgreich eingefügt');
                    // Optional: Textbereich leeren nach erfolgreicher Einfügung
                    document.getElementById('htmlCode').value = '';
                }
            }
        );
    }
})();

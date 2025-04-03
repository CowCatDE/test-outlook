(function() {
    'use strict';

    Office.onReady(function() {
        document.getElementById('insertHtmlButton').onclick = insertHtml;
    });

    function insertHtml() {
        var htmlCode = document.getElementById('htmlCode').value;
        
        // Der Trick: HTML in einen "data:" Link verpacken
        var emailContent = 'Bitte klicken Sie auf den Link, um den HTML-Code anzuzeigen:<br>' +
                          '<a href="data:text/html;charset=utf-8,' + encodeURIComponent(htmlCode) + 
                          '" target="_blank">HTML-Code anzeigen</a>';
        
        // Alternative: HTML-Code als Anhang anbieten
        var attachmentOption = document.getElementById('attachOption');
        if (attachmentOption && attachmentOption.checked) {
            // Diese Funktion wird im Textbereich der E-Mail einen Hinweis hinzufügen,
            // dass der HTML-Code als Anhang beigefügt ist
            emailContent += '<br><br>Der HTML-Code wurde auch als .html Datei angehängt.';
            
            // HTML als Anhang hinzufügen (erfordert erweiterte Berechtigungen)
            var fileName = "html-code.html";
            var fileContent = new Blob([htmlCode], {type: 'text/html'});
            
            // Hinweis: Das Hinzufügen von Anhängen erfordert erweitertes Add-In und entsprechende Berechtigungen
            // Diese Implementierung ist nur konzeptionell
            
            // TODO: Datei als Anhang hinzufügen (erweiterte Implementierung erforderlich)
        }
        
        // In die E-Mail einfügen
        Office.context.mailbox.item.body.setSelectedDataAsync(
            emailContent,
            { coercionType: Office.CoercionType.Html },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Fehler beim Einfügen des Links:', result.error.message);
                } else {
                    console.log('Link erfolgreich eingefügt');
                    document.getElementById('htmlCode').value = '';
                }
            }
        );
    }
})();

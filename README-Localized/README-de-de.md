# <a name="word--javascript-redact-add-in"></a>JavaScript-Redact-Add-In für Word

Erfahren Sie, wie Sie ein Add-In erstellen können, das Text sucht, hervorhebt und bearbeitet.    

## <a name="table-of-contents"></a>Inhalt
* [Änderungsverlauf](#change-history)
* [Wichtiger Hinweis](#important-note)
* [Voraussetzungen](#prerequisites)
* [Konfigurieren des Projekts](#configure-the-project)
* [Ausführen des Projekts](#run-the-project)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

25. Juli 2016:
* Erste Beispielversion

## <a name="important-note"></a>Wichtiger Hinweis

In diesem Beispiel wird Text bearbeitet, indem jeder Buchstabe durch ein Sternchen ersetzt wird.  Das Layout als solches wird nicht beibehalten.  Es wird keine Messung und Berechnung ausgeführt, um die Schriftart, Größe oder Formatierung der Buchstaben zu ermitteln.

## <a name="prerequisites"></a>Voraussetzungen

* Word 2016 für Windows, Build 16.0.6912.1000 oder höher.
* [Node und npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) Sie sollten eine höhere Buildversion verwenden, da bei früheren Buildversionen ein Fehler beim Generieren der Zertifikate auftreten kann.

## <a name="configure-the-project"></a>Konfigurieren des Projekts

Führen Sie folgende Befehle in der Bash-Shell im Stammverzeichnis dieses Projekts aus:

1. Klonen Sie dieses Repository auf ihrem lokalen Computer.
2. ```npm install``` zum Installieren aller Abhängigkeiten in package.json.
3. ```bash gen-cert.sh``` zum Erstellen der für die Ausführung dieses Beispiels erforderlichen Zertifikate. 
* Doppelklicken Sie dann im Repository auf dem lokalen Computer auf „ca.crt“, und wählen Sie **Zertifikat installieren**. 
* Wählen Sie **Lokaler Computer** aus, und wählen Sie **Weiter**, um den Vorgang fortzusetzen. 
* Wählen Sie die Option **Alle Zertifikate in folgendem Speicher speichern**, und wählen Sie dann **Durchsuchen**.  
* Wählen Sie **Vertrauenswürdige Stammzertifizierungsstellen** und dann **OK**. 
* Wählen Sie **Weiter** und dann **Fertig stellen**. Ihre Zertifizierungsstelle wurde nun zum Zertifikatspeicher hinzugefügt.
4. ```npm start``` zum Starten des Knoten-Webserverdiensts.

Jetzt müssen Sie Microsoft Word mitteilen, wo es das Add-In finden kann.

1. Erstellen Sie eine Netzwerkfreigabe oder [geben Sie einen Ordner im Netzwerk frei](https://technet.microsoft.com/en-us/library/cc770880.aspx), und platzieren Sie die [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml)-Manifestdatei darin.
3. Starten Sie Word, und öffnen Sie ein Dokument.
4. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
5. Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
6. Klicken Sie auf **Kataloge vertrauenswürdiger Add-Ins**.
7. Geben Sie in das Feld **Katalog-URL** den Netzwerkpfad zur Ordnerfreigabe an, die die Datei „word-add-in-js-redact-manifest.xml“ enthält, und wählen Sie dann **Katalog hinzufügen**.
8. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und klicken Sie dann auf **OK**.
9. Es wird eine Meldung angezeigt, dass Ihre Einstellungen angewendet werden, wenn Sie Microsoft Office das nächste Mal starten.

## <a name="run-the-project"></a>Ausführen des Projekts

1. Öffnen Sie ein Word-Dokument.
2. Klicken Sie auf der Registerkarte **Einfügen** in Word 2016 auf **Meine-Add-Ins**.
3. Klicken Sie auf die Registerkarte **FREIGEGEBENER ORDNER**.
4. Wählen Sie **Word Redact-Add-In**, und wählen Sie dann **OK**.
5. Wenn Add-In-Befehle von Ihrer Word-Version unterstützt werden, werden Sie in der Benutzeroberfläche darüber informiert, dass das Add-In geladen wurde.

### <a name="ribbon-ui"></a>Menüband-Benutzeroberfläche

Im Menüband:
* Wählen Sie die Registerkarte **Überprüfen** aus, und wählen Sie die Option **Bearbeitungs-Aufgabenbereich anzeigen** aus, um den Aufgabenbereich zu starten.

 > Hinweis: Das Add-In wird in einem Aufgabenbereich geladen, wenn Add-In-Befehle von Ihrer Version von Word nicht unterstützt werden.

### <a name="task-pane-ui"></a>Aufgabenbereich-Benutzeroberfläche

Im Aufgabenbereich können Sie folgende Aktionen ausführen:
* Suchen und Hervorheben von Text in einem Dokument durch Eingeben eines Worts in das Textfeld und Auswählen der Schaltfläche **Suchen und Hervorheben**.
  
> HINWEIS:  Bei der Suche wird zwischen Groß- und Kleinschreibung unterschieden.  Drücken Sie STRG+Z auf der Tastatur, um eine Aktion rückgängig zu machen.

* Bearbeiten Sie Text in einem Dokument, indem Sie ein Wort in das Textfeld eingeben und die Schaltfläche **Redact** auswählen.
  
> HINWEIS:  Bei Redact wird zwischen Groß-und Kleinschreibung unterschieden.   

* Bearbeiten Sie ausgewählten Text in einem Dokument, indem Sie die Schaltfläche **Ausgewählten Text bearbeiten** auswählen.
  
> HINWEIS:  Bei Redact wird zwischen Groß-und Kleinschreibung unterschieden.       
  
### <a name="add-in-commands"></a>Add-In-Befehle

In diesem Beispiel wird gezeigt, wie Sie Add-In-Befehle im Menüband erstellen. Die Add-In-Befehldeklarationen sind in der Datei „ word-add-in-js-redact-manifest.xml“ enthalten. 

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des Word Redact-Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Allgemeine Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen mit [office-js] und [API] markiert sind.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* 

  [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs – Startprojekte und Codebeispiele](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Alle Rechte vorbehalten.



In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).

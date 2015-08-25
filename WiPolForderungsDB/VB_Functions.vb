Rem CL034: Neue Funktion zur Auswahl und Speicherung des Vorlagenordners
Private Sub bVorlagenordner_Click()
    
    Dim fd As FileDialog
    Dim pfad As Variant
    Dim db As Database
    Set db = CurrentDb

    Rem Erzeuge und konfiguriere Folder Picker Dialog Objekt
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.AllowMultiSelect = False
    fd.Title = "Vorlagenordner auswählen"
    fd.InitialFileName = DLookup("Vorlagenordner", "tKonfiguration", "KonfigurationID = 1")
    
    Rem Zeige File Dialog an und verarbeite Auswahl
    If fd.Show = True Then
        Rem Wenn Benutzer Auswahl getroffen hat, speichere Pfad in Datenbank
        For Each pfad In fd.SelectedItems
            db.Execute ("UPDATE tKonfiguration SET Vorlagenordner = '" & pfad & "' WHERE KonfigurationID = 1")
            db.Close
            
            aktualisiereVorlage
        Next
    Else
        Rem Cancel -> Auswahl nicht erfolgreich
        Exit Sub
    End If
    
End Sub


Rem CL036: Neue Funktion zum Einlesen der Auswahlmöglichkeiten für Vorlagendatei aus Vorlagenordner
Private Sub aktualisiereVorlage()
    
    Rem CL039: Deklarationen zur Vermeidung von Kompilierfehlern in manchen Umgebungen
    Dim newSource, FS, Folder, file
    
    Me.Vorlage.RowSource = ""
    Me.Vorlage.Value = ""
    newSource = ""
    
    On Error GoTo FolderNotFound ' für den Fall, dass der eingestellte Ordner nicht gefunden wird
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set Folder = FS.GetFolder(DLookup("Vorlagenordner", "tKonfiguration", "KonfigurationID = 1"))
    For Each file In Folder.Files
        If file.Name Like "*.doc*" Then
            newSource = newSource & "'" & file.Name & "';"
        End If
    Next
    
FolderNotFound:
    On Error GoTo 0 ' ab hier werden Fehler wieder normal geworfen
    
    Me.Vorlage.RowSource = newSource
    Me.Vorlage.Requery
    If Me.Vorlage.ListCount > 0 Then
        Me.Vorlage.SetFocus
        Me.Vorlage.ListIndex = 0
    Else
        MsgBox "Für den Word-Export muss eine Vorlagendatei ausgewählt werden!"
    End If
    
End Sub

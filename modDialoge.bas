Attribute VB_Name = "modDialoge"
Option Explicit

'Beschreibung
'------------
    'Mit der Show-Methode des Application-Objekts
    'lassen sich �ber die Application.Dialogs-Eigenschaft
    'gezielt die einzelnen Excel-Dialoge aufrufen.
    'Eine komplette Auflistung der vorhandenen Dialoge
    'wird nach Eintippen von Application.Dialogs( nach der
    '�ffnenden Klammer angezeigt (sofern in den Optionen
    'im VBA-Editor unter "Extras" der Haken bei "Elemente
    'automatisch auflisten" gesetzt ist) bzw. durch
    'Strg+Leertaste nach der Klammer, falls die Auflistung
    'nicht freiwillig aufpoppt.
    
'Code
'----
Sub DialogAufrufen() 'Hier eine kleine Auswahl an m�glichen Dialogfenstern

        'Zellen formatieren - Schrift
            Application.Dialogs(xlDialogActiveCellFont).Show
        'Zellen formatieren - Schutz
            Application.Dialogs(xlDialogCellProtection).Show
        'Blatt umbenennen
            Application.Dialogs(xlDialogWorkbookName).Show
        'Struktur und Fenster sch�tzen
            Application.Dialogs(xlDialogWorkbookProtect).Show
        'Add-Ins
            Application.Dialogs(xlDialogAddinManager).Show
            
End Sub

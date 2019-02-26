Attribute VB_Name = "Modul1"

'(C) Michael G�mmer, 24.06.2016
'Einf�gen von Bildern in eine bestimmte Spalte der Exeltabelle.
'nach dem Dateinamen der in einer anderen Spalte steht
  
'Prozedur zum Einf�gen der Bilder
Sub Bilder_einf�gen()
 
  'Verzeichnis in dem Die Bilder liegen
  Dim Pfad, picPfad, RW_Start_S As String
  Dim i, range_A As Integer
  Dim CL_pic, CL_pname, RW_Start As Integer
  
  RW_Start = 6 'Reihe ab der die Tabelle Startet
  CL_pname = 1 'Spalt in dem der Dateiname f�r die Bilder steht (in der Regel = Artikelnummer)
  CL_pic = 2   'Spalt in die die Bilder eingef�gt werden
  On Error Resume Next  'Fehler Abfangen

   '### Pfad zu den Produktbildern - Optionen (1) - (2) k�nnen durch L�schen des Hochkommas aktiviert werden.
  
  '(1) Pfad zum aktuellen Verzeichnis - in dem sich die Excel Tabelle befindet. Hierzu mussen alle Bilder im selben Verzeichnis sein
  'Pfad = ThisWorkbook.Path & "\"'
 
  '(2) Pfad zum Quasar3 Clipart-Verzeichnis
  'Pfad = "R:\Quasar3\Images\Artikel\"
  
  '(3) Pfad zu den Buffetplaner Grafiken
  Pfad = "P:\FRICH Buffetplaner\"

  Rows(RW_Start & ":10000").Select    'Auswahl aller Zeilen von 4 bis 10000
  Selection.RowHeight = 13.2          'Zeilenh�he auf 13,2 (Standard)
    
  range_A = Range("A10000").End(xlUp).Row
        
  Rows(RW_Start & ":" & range_A).Select   'Auswahl aller Zeilen von 4 bis Ende der Tabelle
  Selection.RowHeight = 86      'Zeilenh�he auf 86 (passend zu den Quasar3 Cliparts
  Columns("B:B").Select         'Auswahl der Bild-Spalte
  Selection.ColumnWidth = 15    'Breite auf 15
   
 
  For i = 4 To range_A                                ' Z�hler bis zum Ende der Tabelle
    Cells(i, CL_pic).Select                           ' Auswahl der Zelle, in die das Bild soll
    Cells(i, CL_pic).Activate                         ' Aktivierung der Zelle
    picPfad = Pfad & Cells(i, CL_pname) & ".gif"      ' Bilddateipfad wird aus dem Wert, der in Spalte 4 steht generiert
        
    If Dir(picPfad) <> "" Then                 ' Wenn die Bilddatei existiert, Bild einf�gen
                                               '(Bildpfad, als Verkn�pfung ?, Bild mit Datei speichern?, Position und Ma�e der aktuellen Zelle)
       ActiveSheet.Shapes.AddPicture(picPfad, False, True, Selection.Left, Selection.Top, ActiveCell.Width, ActiveCell.Height).Select
    Else                                       ' Wenn die Bilddatei nicht existiert, dann leeres Bild einf�gen
       'ActiveSheet.Shapes.AddPicture(Pfad & "nothing.jpg", False, True, Selection.Left, Selection.Top, ActiveCell.Width, ActiveCell.Height).Select
       ActiveCell.FormulaR1C1 = "kein Bild vorhanden"
    End If
  Next i
End Sub

'Prozedur zum L�schen der Bilder


Sub Bilder_loeschen()
    ActiveSheet.Pictures.Delete
     Rows("5:10000").Select    'Auswahl aller Zeilen von 4 bis 10000
  Selection.RowHeight = 13.2   'Zeilenh�he auf 13,2 (Standard)
End Sub


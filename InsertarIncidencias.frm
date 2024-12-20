VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertarIncidencias 
   Caption         =   "Añadir incidencia"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "InsertarIncidencias.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "InsertarIncidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public destFolder As String
    Public ws As Worksheet
    Public MyTable As ListObject
    Public tblRow As Range
    Public NumeroInc As String
Private Sub CommandButton1_Click()
    Dim fd As FileDialog
    If ComboBox1.value = "" Then
        MsgBox ("Especifica el tipo de incidencia primero.")
    Else
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        Dim vrtSelectedItem As Variant
        Dim fso As Object
        Dim fileName As String
        
        destFolder = "\Enlaces\" & ComboBox1.value & "\" & Format(Date, "yyyy") & "\" & TextBox2.value & "\"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not fso.FolderExists(destFolder) Then
            fso.CreateFolder destFolder
        End If
        
        With fd
            If .Show = -1 Then
                For Each vrtSelectedItem In .SelectedItems
                    fileName = fso.GetFileName(vrtSelectedItem)
                    fso.CopyFile vrtSelectedItem, destFolder & fileName
                Next vrtSelectedItem
                MsgBox "Los archivos se han guardado correctamente en " & destFolder
            End If
        End With
        
        Set fd = Nothing
        Set fso = Nothing
    End If
End Sub

Private Sub CommandButton2_Click()
    Set ws = Sheets("INC-ABIERTAS")
    Set MyTable = ws.ListObjects(1)
    Set tblRow = MyTable.ListRows.Add.Range
    Dim Hyperlink As String

    If destFolder <> "" Then
        Hyperlink = "=HYPERLINK(""" & destFolder & """, """ & TextBox2.value & """)"
    Else
        Hyperlink = ""
    End If

    With tblRow
        .Cells(8) = TextBox2.value
        .Cells(9) = ComboBox1.value
        .Cells(10) = TextBox1.value
        .Cells(12) = TextBox3.value
        .Cells(13) = TextBox4.value
        .Cells(14) = TextBox5.value
        .Cells(15) = TextBox23.value
        .Cells(16) = TextBox9.value
        .Cells(17) = ComboBox2.value
        .Cells(18) = TextBox24.value
        .Cells(19) = ComboBox3.value
        .Cells(20) = TextBox13.value
        .Cells(21) = TextBox25.value
        .Cells(22) = TextBox26.value

        ' archivos
        .Cells(23) = Hyperlink
        ' sev
        .Cells(24) = TextBox20.value
        ' occ
        .Cells(25) = TextBox21.value
        ' det
        .Cells(26) = TextBox22.value
    End With

    MsgBox "Incidencia añadida correctamente."
    Dim EnviarEmail As String

    EnviarEmail = MsgBox("¿Te gustaría enviar un correo con estos datos?", vbYesNo + vbQuestion, "¿Enviar email?")

    If EnviarEmail = vbYes Then

        Dim OutlookApp As Object
        Dim OutlookMail As Object
        Dim ws2 As Worksheet
        Dim CodeToMatch As String
        Dim LastRow As Long
        Dim i As Long
        Dim emailBody As String
        Dim rowData As String
        Dim hyperlinkCell As Range
        Dim folderLink As String

        Set ws2 = ThisWorkbook.Sheets("INC-ABIERTAS")

        CodeToMatch = TextBox2.value

        LastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row

        For i = 1 To LastRow
            If ws2.Cells(i, 8).value = CodeToMatch Then
                rowData = ""

                Dim columnsToInclude As Variant
                columnsToInclude = Array(8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 19, 21, 23)

                emailBody = "<html><body><br><br><br><h2>Información de la incidencia " & CodeToMatch & "</h2><br><br><br>"
                emailBody = emailBody & "<table border='1' cellpadding='5' cellspacing='0'>"
                emailBody = emailBody & "<tr><th>Campo</th><th>Valor</th></tr>"

                For Each col In columnsToInclude
                    If ws2.Cells(i, col).value <> "" Then
                        If col = 23 Then
                            Set hyperlinkCell = ws2.Cells(i, col)

                            If hyperlinkCell.HasFormula Then
                                On Error Resume Next
                                folderLink = Mid(hyperlinkCell.Formula, InStr(1, hyperlinkCell.Formula, """") + 1, _
                                                  InStr(InStr(1, hyperlinkCell.Formula, """") + 1, hyperlinkCell.Formula, """") - _
                                                  InStr(1, hyperlinkCell.Formula, """") - 1)
                                On Error GoTo 0

                                If folderLink <> "" Then
                                    rowData = "<tr><td>" & ws2.Cells(1, col).value & "</td><td><a href='" & folderLink & "'>" & CodeToMatch & "</a></td></tr>"
                                    emailBody = emailBody & rowData
                                End If
                            End If
                            
                        ElseIf col = 14 Then
                            rowData = "<tr><td>" & ws2.Cells(1, col).value & "</td><td style='color: red;'>" & ws2.Cells(i, col).value & "</td></tr>"
                            emailBody = emailBody & rowData
                        Else
                            rowData = "<tr><td>" & ws2.Cells(1, col).value & "</td><td>" & ws2.Cells(i, col).value & "</td></tr>"
                            emailBody = emailBody & rowData
                        End If
                    End If
                Next col

                emailBody = emailBody & "</table></body></html>"

                Exit For
            End If
        Next i

        If emailBody <> "" Then
            Set OutlookApp = CreateObject("Outlook.Application")
            Set OutlookMail = OutlookApp.CreateItem(0)

            With OutlookMail
                .Subject = "Incidencia " & CodeToMatch
                .BodyFormat = olFormatHTML
                .HTMLBody = emailBody
                .To = ""
                .Display
            End With
        Else
            MsgBox "Número de incidencia " & CodeToMatch & " no encontrado.", vbExclamation
        End If
    End If
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

Private Sub TextBox23_Enter()
TextBox23.value = Format(CalendarForm.GetDate, "dd/mm/yyyy")
End Sub

Private Sub TextBox24_Enter()
TextBox24.value = Format(CalendarForm.GetDate, "dd/mm/yyyy")
End Sub

Private Sub TextBox25_Enter()
TextBox25.value = Format(CalendarForm.GetDate, "dd/mm/yyyy")
End Sub

Private Sub TextBox26_Enter()
TextBox26.value = Format(CalendarForm.GetDate, "dd/mm/yyyy")
End Sub
Private Sub UserForm_Initialize()
    Set ws = Sheets("INC-ABIERTAS")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    With Me.TextBox2
        Dim MyDate
        Dim FilaNumeros
        Dim Fecha
        Dim Iterador As Integer
        Dim IteradorFormato
    
        Iterador = 1
        IteradorFormato = Format(Iterador, "00")
        MyDate = Date
        Fecha = Format(MyDate, "yy.mm.dd")
        NumeroInc = Fecha & IteradorFormato
    
    Do While Application.WorksheetFunction.CountIf(ws.Range("H:H"), NumeroInc) <> 0
        Iterador = Iterador + 1
        IteradorFormato = Format(Iterador, "00")
        NumeroInc = Fecha + IteradorFormato
    Loop


        TextBox2.value = NumeroInc
End With

    With Me.ComboBox1

    .DropButtonStyle = fmDropButtonStyleArrow
    .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    .Style = fmStyleDropDownCombo
    
    .AddItem "Proceso"
    .AddItem "Cliente"

End With
    With Me.ComboBox2
    Dim Filas3

    Filas3 = Sheets("DATOS").Cells(Sheets("DATOS").Rows.Count, "D").End(xlUp).Row
    ActiveWorkbook.Names.Add Name:="Modelos", RefersTo:=Sheets("DATOS").Range("D2:D" & Filas3)
    Me.ComboBox2.RowSource = "Modelos"
End With
    With Me.ComboBox3
    Dim Filas2

    Filas2 = Sheets("DATOS").Cells(Sheets("DATOS").Rows.Count, "F").End(xlUp).Row
    ActiveWorkbook.Names.Add Name:="Clientes", RefersTo:=Sheets("DATOS").Range("F2:F" & Filas2)
    Me.ComboBox3.RowSource = "Clientes"
    
End With

End Sub

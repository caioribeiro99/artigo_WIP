Attribute VB_Name = "ExcelExportMod"
Option Compare Database
Option Explicit

Function PrepXL()
'macro utilizada na distribui��o do arquivo em excel (relat�rio de produ��o)
''m�dulo em desuso devido a sistem�tica de utiliza��o da planilha

Dim xlFile As String
Dim xlFolder As String
Dim NewxlFile As String
Dim xlApp As Excel.Application
Dim xlWbk As Excel.Workbook
Dim cn As Object
Dim qry As Object

    xlFile = "Relat�rio de Produ��o.xlsx"
    xlFolder = CurrentProject.Path & "\"
    xlFile = xlFolder & xlFile
    NewxlFile = InputBox("Insira o nome do novo arquivo", "WIP MFPA")
    If Nz(NewxlFile, "") = "" Then
      Exit Function
    End If
    If Dir(xlFile) = "" Then
      MsgBox "O Arquivo" & xlFile & "n�o foi encontrado!", vbOKOnly, "WIP MFPA"
      Exit Function
    End If
    NewxlFile = xlFolder & NewxlFile & ".xlsx"
    If Dir(NewxlFile) <> "" Then
      Kill (NewxlFile)
    End If
    FileCopy xlFile, NewxlFile
    Set xlApp = Excel.Application
    Set xlWbk = xlApp.Workbooks.Open(NewxlFile)
    xlWbk.EnableConnections
    DoEvents
    xlWbk.RefreshAll
    DoEvents
    xlWbk.Save
    DoEvents
    On Error Resume Next
    For Each cn In xlWbk.Connections
        cn.Delete
    Next cn
    For Each qry In xlWbk.Queries
        qry.Delete
    Next qry
    xlWbk.Save
    DoEvents
    xlWbk.Close
    DoEvents
    Set xlApp = Nothing
    Set xlWbk = Nothing
    MsgBox "Relat�rio de Produ��o exportado com sucesso!", vbOKOnly, "WIP MFPA"
End Function

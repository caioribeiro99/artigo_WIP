Attribute VB_Name = "ExibicaoMod"
Option Compare Database
Option Explicit

Function Startup()
'macro para carregar a tela
    
    EventExec = False
    EventExec2 = False
    SelectOPA = True
    HideVal = True
    OpenFile = True
    Modo_Menu
    'CheckVersion
    DoCmd.OpenForm "Login", , , , , acDialog
End Function

Sub Modo_Menu()
'macro para habilitar o modo de visualização de apontamento (desabilitando navegação e edição)

    If HideVal = True Then
        DoCmd.NavigateTo "acNavigationCategoryObjectType"
        DoCmd.RunCommand acCmdWindowHide
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
        HideVal = False
    Else
        DoCmd.SelectObject acTable, , True
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
        HideVal = True
    End If
End Sub

Function Modo_Edicao()
'macro para restaurar a tela para o modo de edição
On Error Resume Next
    HideVal = False
    Modo_Menu
    DoCmd.Close acForm, "Painel_do_Gestor"
    DoCmd.Close acForm, "Login"
    DoCmd.Close acForm, "Apontamento"
End Function

Function CheckVersion()
'função para verificar se existem atualizações disponíveis no sistema

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String

    On Error GoTo ErrHandler
    Set db = CurrentDb
    Set Backlog = db.OpenRecordset("ErrorBacklog")
    FileName = CurrentProject.Name
    LocalObject = CurrentObjectName
    
    Dim SrvrVer As Integer
    Dim StnVer As Integer
    Dim VersionName As String
    Dim rst As Recordset
    Set db = CurrentDb
    Set rst = db.OpenRecordset("ServerVersion")
    rst.MoveFirst
    SrvrVer = rst!Version
    VersionName = rst!VersionName
    rst.Close
    Set rst = db.OpenRecordset("StationVersion")
    StnVer = rst!Version
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    If SrvrVer > StnVer Then
       MsgBox "O sistema foi atualizado para a versão " & VersionName & ". Atualize o sistema para poder continuar!", vbOKOnly, "WIP MFPA"
       DoCmd.Quit
       Exit Function
    End If
Exit Function
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!Código = Err.Number
    Backlog!Descrição = Err.Description
    Backlog!Local = FileName & " / " & LocalObject
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Function

    

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Encerradas_Visualização"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim SrchVal As String
Dim SrchCrit As String
Dim LastFld As String
Dim SrchVal2 As String
Dim SrchCrit2 As String


Private Sub Form_Current()
'macro para sistematizar o sincronismo dos registros entre o sub-formulário e o formulário principal

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String

On Error GoTo ErrHandler
If EventExec2 Then Exit Sub

Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
On Error Resume Next
LocalControl = Me.ActiveControl.Name
On Error GoTo ErrHandler

Dim ctl As Control
Dim fldName As String
Dim rst As Recordset

If OpenFile Then GoTo Final

If SelectOPA = True Then
'    Me.Parent.OPA_Atual = Me.OPA
'    Me.Parent.Requery
'    Me.Parent.BP.SetFocus
    If Nz(Me.OPA, "") <> "" Then
        MsgBox "Exibindo último apontamento da OPA: " & Me.OPA, vbOKOnly, "WIP MFPA"
    End If
Else
    Set ctl = Screen.PreviousControl
    fldName = "OPA"
    SrchVal2 = PrevOPAVal
    SrchCrit2 = "[" & fldName & "] = " & SrchVal2
    Set rst = Me.RecordsetClone
    rst.FindFirst SrchCrit2
    If rst.NoMatch Then
        EventExec2 = False
        Exit Sub
    Else
        EventExec2 = True
        Me.Bookmark = rst.Bookmark
        GoTo Final
    End If
'    Me.Parent.OPA_Atual = PrevOPAVal
End If
EventExec2 = False
DoEvents
Exit Sub
Final:
EventExec2 = False
OpenFile = False
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!Código = Err.Number
    Backlog!Descrição = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para desabilitar as teclas na visualização do subformulário

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String

Dim ctl As Control
Dim fldName As String
Dim rst As Recordset

On Error GoTo ErrHandler
Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
On Error Resume Next
LocalControl = Me.ActiveControl.Name
On Error GoTo ErrHandler

Select Case KeyCode
    Case vbKeyRight, vbKeyLeft
    Case vbKeyEnd
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToLast
   Case vbKeyHome
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToFirst
   Case vbKeyUp
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToPrevious
   Case vbKeyDown
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToNext
    Case 48 To 57, 65 To 90, 97 To 105 'todas as letras e números
'        Set ctl = Screen.ActiveControl
        fldName = "OPA"
'        If fldName <> LastFld Then
'            SrchVal = ""
'        End If
'        LastFld = fldName
        SrchVal = SrchVal & Chr(KeyCode)
        If SrchVal <> "" Then
            Me.OPA_Pesquisa.Visible = True
            Me.OPA_Pesquisa_Ícone.Visible = True
        Else
            Me.OPA_Pesquisa.Visible = False
            Me.OPA_Pesquisa_Ícone.Visible = False
        End If
        SrchCrit = "[" & fldName & "] Like '" & SrchVal & "*'"
        Me.OPA_Pesquisa.Value = SrchVal
        KeyCode = 0
    Case Is = 13 'tecla enter
        Set rst = Me.RecordsetClone
        rst.FindFirst SrchCrit
        If rst.NoMatch Then
            MsgBox "Registro não encontrado!", vbOKOnly, "WIP MFPA"
            SrchVal = ""
            rst.Close
            Me.OPA_Pesquisa.Visible = False
            Me.OPA_Pesquisa_Ícone.Visible = False
            Me.OPA_Pesquisa.Value = ""
            KeyCode = 0
            Exit Sub
        End If
        Me.Bookmark = rst.Bookmark
        rst.Close
        SrchVal = ""
        Me.OPA_Pesquisa.Visible = False
        Me.OPA_Pesquisa_Ícone.Visible = False
        Me.OPA_Pesquisa.Value = ""
        KeyCode = 0
    Case Is = 27 'tecla esc
        SrchVal = ""
        Me.OPA_Pesquisa.Visible = False
        Me.OPA_Pesquisa_Ícone.Visible = False
        Me.OPA_Pesquisa.Value = ""
        KeyCode = 0
    Case Else
        Me.OPA_Pesquisa.Visible = False
        Me.OPA_Pesquisa_Ícone.Visible = False
        Me.OPA_Pesquisa.Value = ""
        KeyCode = 0
    End Select
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!Código = Err.Number
    Backlog!Descrição = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para sincronizar a seleção de registro do formulário principal e secundário no momento de abertura do aplicativo
    On Error Resume Next
    DoCmd.RunCommand acCmdRecordsGoToNew
End Sub

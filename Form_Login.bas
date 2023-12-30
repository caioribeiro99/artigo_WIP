VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ExitBtn_Click()
'macro para sair da tela de login e fechar o programa

    If MsgBox("Você realmente deseja sair do programa?", vbYesNo, "WIP MFPA") = vbNo Then
        Exit Sub
    End If
    DoCmd.Quit
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para carregar o formulário de login

    Me.UserName = ""
    Me.Pword = ""
    Me.UserName.SetFocus
End Sub

Function TestEntry()
'macro para realizar a validação do usuário e senha

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
LocalControl = Me.ActiveControl.Name

    If Nz(Me.UserName, "") = "" Then
        MsgBox "Campo de usuário está vazio!", vbCritical, "WIP MFPA"
        TestEntry = False
        Exit Function
    End If
    If Nz(Me.Pword, "") = "" Then
        MsgBox "Campo de senha está vazio!", vbCritical, "WIP MFPA"
        Me.Pword.SetFocus
        TestEntry = False
        Exit Function
    End If
    If Me.Pword <> Me.UserName.Column(1) Then
        MsgBox "Senha incorreta!", vbCritical, "WIP MFPA"
        Me.Pword = ""
        Me.Pword.SetFocus
        TestEntry = False
        Exit Function
    End If
    TestEntry = True
Exit Function
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
End Function

Private Sub LoginBtn_Click()
'macro para realizar login

    If TestEntry = True Then
        DoCmd.Close acForm, "Login"
        DoCmd.OpenForm "Painel_do_Gestor", , , , , acDialog
    End If
End Sub

Private Sub Pword_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizas funções da tecla enter e esc
    Select Case KeyCode
        Case Is = 13
            Me.Dirty = False
            LoginBtn_Click
        Case Is = 27
            ExitBtn_Click
    End Select
End Sub

Private Sub UserName_AfterUpdate()
'macro para sistematizer ordem de preenchimento dos campos
Me.Pword.SetFocus
End Sub

Private Sub UserName_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizas funções da tecla enter e esc
    Select Case KeyCode
        Case Is = 13
            Me.Pword.SetFocus
        Case Is = 27
            ExitBtn_Click
    End Select
End Sub

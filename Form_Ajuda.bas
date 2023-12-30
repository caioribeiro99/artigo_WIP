VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Ajuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
'macro para inicializar

Me.Voltar.SetFocus
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para verificar a versão do sistema

Dim db As DAO.Database
Dim rst As Recordset
Dim Version As String

Set db = CurrentDb
Set rst = db.OpenRecordset("StationVersion")

rst.MoveFirst
Version = rst!VersionName
Me.Versao_Sistema = Version

Set db = Nothing
Set rst = Nothing
End Sub

Private Sub Voltar_Click()
'macro para retornar ao formulário de apontamento

    DoCmd.Close acForm, "Ajuda"
    OpenFile = True
    DoCmd.OpenForm "Apontamento", , , , , acDialog
End Sub

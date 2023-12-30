VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Apontamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Ajuda_Click()
'macro para abrir o formulário de ajuda

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

    If Me.Dirty = True Or SelectOPA = False Then
        MsgBox "Houveram alterações não salvas no registro anterior! Por favor, SALVE ou DESFAÇA as alterações para poder " & _
        "continuar. ", vbOKOnly, "WIP MFPA"
        SelectOPA = False
    Else
        DoCmd.Close acForm, "Apontamento"
        DoCmd.OpenForm "Ajuda", , , , , acDialog
    End If
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

Private Sub AN_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizar função específicas nas teclas
    If KeyCode = 27 Then
        Desfazer_Click
    End If
End Sub

Private Sub APAGAR_Click()
'macro para apagar o registro atual ou limpar os campos de apontamento

Dim StrSQL As String
Dim Operação As String
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

Me.Ícone_Apagar.SetFocus
Operação = Nz(Me.Operação.Value, "")

    If MsgBox("Você realmente deseja APAGAR os dados desse registro? Essa ação não poderá ser desfeita", vbYesNo, "WIP MFPA") = vbYes Then
        Select Case Operação
            Case Is = "ENROLAMENTO", "FORMADORA", "ESTAMPAGEM"
                StrSQL = "DELETE Histórico.*, Histórico.OPA, Histórico.Operação FROM Histórico " & _
                "WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & Me.Operação & "'"
                DoCmd.RunSQL StrSQL
                DoCmd.RunCommand acCmdDeleteRecord
                Me.Undo
                Me.Refresh
            Case Else
                StrSQL = "DELETE Histórico.*, Histórico.OPA, Histórico.Operação FROM Histórico " & _
                "WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & Me.Operação & "'"
                DoCmd.RunSQL StrSQL
                DoCmd.RunCommand acCmdDeleteRecord
                Me.Undo
                Me.Programa_Rótulo.Visible = False
                Me.Programa.Visible = False
                Me.Repasse_Rótulo.Visible = False
                Me.Repasse.Visible = False
                Me.Refresh
        End Select
    End If
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

Private Sub BP_AfterUpdate()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
On Error Resume Next
    Me.NP.SetFocus
End Sub

Private Sub BP_GotFocus()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    If Me.Operação = "PRETEJAMENTO" Then
        Me.Programa_Rótulo.Visible = True
        Me.Programa.Visible = True
        Me.Repasse_Rótulo.Visible = True
        Me.Repasse.Visible = True
    Else
        Me.Programa_Rótulo.Visible = False
        Me.Programa.Visible = False
        Me.Repasse_Rótulo.Visible = False
        Me.Repasse.Visible = False
    End If
End Sub

Private Sub BP_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para adicionar funções específicas para as teclas
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub BP_LostFocus()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    If Me.Operação = "PRETEJAMENTO" Then
        Me.Programa_Rótulo.Visible = True
        Me.Programa.Visible = True
        Me.Repasse_Rótulo.Visible = True
        Me.Repasse.Visible = True
    Else
        Me.Programa_Rótulo.Visible = False
        Me.Programa.Visible = False
        Me.Repasse_Rótulo.Visible = False
        Me.Repasse.Visible = False
    End If
End Sub

Private Sub Centro_de_Custo_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.NP.SetFocus
End Sub

Private Sub Data_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
If SelectOPA = True Then
    Me.Recalc
    Me.OPA.SetFocus
End If
End Sub

Private Sub Desfazer_Click()
'macro para desfazer as alterações do registro mostrado no formulário principal

On Error GoTo ErrHandler

Dim Operação As String
Dim dattime As Date
Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String

Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
LocalControl = Me.ActiveControl.Name

Operação = Nz(Me.Operação.Value, "")
dattime = DateAdd("s", 0.5, Now)

EventExec = True
EventExec2 = True
SelectOPA = True

    If Me.NewRecord = False Then
        Select Case Operação
            Case Is = "PRETEJAMENTO"
                If Me.Dirty = True Then
                    Me.Undo
                    Do
                        DoEvents
                    Loop Until Now >= dattime
                    MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
                Else
                    Me.OPA.Value = PrevOPAVal
                    Me.AN.Value = PrevANVal
                    Me.BP.Value = PrevBPVal
                    Me.NP.Value = PrevNPVal
                    Me.Pirâmide.Value = PrevPrmdVal
                    Me.Parcial.Value = PrevParcVal
                    Me.Programa.Value = PrevProgVal
                    Me.Repasse.Value = PrevRepVal
                    Me.Produzido.Value = PrevProdVal
                    Me.Refugado.Value = PrevRefVal
                    Me.Dirty = False
                    Me.BP.SetFocus
                    Do
                        DoEvents
                    Loop Until Now >= dattime
                    MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
                End If
            Case Else
                If Me.Dirty = True Then
                    Me.Undo
                    Do
                        DoEvents
                    Loop Until Now >= dattime
                    MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
                Else
                    If Nz(PrevANVal, "") = "" And Nz(PrevOPAVal, 0) = 0 Then
                        DoCmd.RunCommand acCmdDeleteRecord
                        MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
                    Else
                        Me.OPA.Value = PrevOPAVal
                        Me.AN.Value = PrevANVal
                        Me.BP.Value = PrevBPVal
                        Me.NP.Value = PrevNPVal
                        Me.Pirâmide.Value = PrevPrmdVal
                        Me.Parcial.Value = PrevParcVal
                        Me.Produzido.Value = PrevProdVal
                        Me.Refugado.Value = PrevRefVal
                        Me.Dirty = False
                        Me.BP.SetFocus
                        Do
                            DoEvents
                        Loop Until Now >= dattime
                        MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
                    End If
                End If
        End Select
    Else
        If Me.Dirty = True Then
            Me.Undo
            Do
                DoEvents
            Loop Until Now >= dattime
            MsgBox "Alterações descartadas!", vbOKOnly, "WIP MFPA"
        End If
    End If
EventExec = False
EventExec2 = False
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

Private Sub Form_BeforeUpdate(Cancel As Integer)
'macro para impedir a troca de registros com alterações não salvas

If EventExec Then Exit Sub
On Error Resume Next

    SelectOPA = True
    EventExec = True
        If Me.NewRecord = True Then
            If MsgBox("Se você sair agora o novo registro será perdido. Deseja continuar?", vbYesNo, "WIP MFPA") = vbYes Then
                Me.Undo
                SelectOPA = True
            Else
                SelectOPA = False
                Exit Sub
            End If
        End If
        If Not Me.NewRecord And Me.Dirty = True Then
            If MsgBox("Se você sair agora as alterações serão perdidas. Deseja continuar?", vbYesNo, "WIP MFPA") = vbYes Then
                Desfazer_Click
            Else
                SelectOPA = False
                Exit Sub
            End If
        End If
    EventExec = False
End Sub

Private Sub Form_Close()
'macro para apagar registros incompletos na saída forçada

If Me.NewRecord Then
    Me.Undo
End If
If Me.Dirty Then
    Me.Undo
End If
End Sub

Private Sub Form_Current()
'macro para guardar os valores dos campos do registro atual (salvamento e restauração)

PrevOPAVal = Nz(Me.OPA.Value, 0)
PrevANVal = Nz(Me.AN.Value, "")
PrevBPVal = Nz(Me.BP.Value, 0)
PrevNPVal = Nz(Me.NP.Value, 0)
PrevProdVal = Nz(Me.Produzido.Value, 0)
PrevRefVal = Nz(Me.Refugado.Value, 0)
PrevPrmdVal = Nz(Me.Pirâmide.Value, 0)
PrevParcVal = Nz(Me.Parcial.Value, 0)
PrevRepVal = Nz(Me.Repasse.Value, 0)
PrevProgVal = Nz(Me.Programa.Value, "")
PrevOperVal = Nz(Me.Operação.Value, "")
End Sub

Private Sub Form_Load()
'macro para sincronizar a ordem dos registros do formulário com o banco de dados e atualizar a versão
On Error Resume Next
Me.BP.SetFocus
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para sincronizar a ordem dos registros do formulário com o banco de dados

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

On Error Resume Next
DoCmd.RunCommand acCmdRecordsGoToNew
Me.BP.SetFocus
End Sub

Private Sub Hora_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
If SelectOPA = True Then
    Me.Recalc
    Me.OPA.SetFocus
End If
End Sub

Private Sub Ícone_Apagar_Click()
'macro para chamar a subrotina de apagar registro do botão associado ao ícone
    APAGAR_Click
End Sub

Private Sub Ícone_Sair_Click()
'macro para chamar a subrotina de sair do aplicativo do botão associado ao ícone
    SAIR_Click
End Sub

Private Sub Ícone_Salvar_Click()
'macro para chamar a subrotina do botão de salvamento de registros do botão associado ao ícone
    SALVAR_Click
End Sub

Private Sub LogoMAHLE_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    On Error Resume Next
    Me!OPA.SetFocus
End Sub

Private Sub Máquina_Enter()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.Parcial.SetFocus
End Sub

Private Sub Nome_do_Colaborador_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.NP.SetFocus
End Sub

Private Sub Nome_do_Colaborador_GotFocus()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.OPA.SetFocus
End Sub

Private Sub Novo_Click()
'macro para funcionamento do botão de novo registro do navegador de registros

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

    If Me.Dirty = True Or SelectOPA = False Then
        MsgBox "Houveram alterações não salvas no registro anterior! Por favor, SALVE ou DESFAÇA as alterações para poder " & _
        "continuar. ", vbOKOnly, "WIP MFPA"
        SelectOPA = False
    Else
        SelectOPA = True
        DoCmd.RunCommand acCmdRecordsGoToNew
        Me.Repasse.Visible = False
        Me.Programa.Visible = False
        Me.Repasse_Rótulo.Visible = False
        Me.Programa_Rótulo.Visible = False
    End If
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Is = 2046
            Resume Next
        Case Else
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
    End Select
End Sub

Private Sub NP_AfterUpdate()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.OPA.SetFocus
End Sub

Private Sub NP_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para adicionar funções específicas para as teclas
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub OPA_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizar a ordem de tabulação e adicionar funções específicas às teclas de navegação entre controles

    If KeyCode = 9 And Shift = 1 Then
        KeyCode = 0
        Me.NP.SetFocus
    End If
    If KeyCode = 38 Or KeyCode = 37 Then
        KeyCode = 0
        Me.NP.SetFocus
    End If
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub Operação_Click()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.BP.SetFocus
End Sub

Private Sub Operação_Enter()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    If Me.Programa.Visible = True And Me.Repasse.Visible = True Then
        Me.Programa.SetFocus
    Else
        Me.Parcial.SetFocus
    End If
End Sub

Private Sub Operação_GotFocus()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    If Me.Operação = "PRETEJAMENTO" Then
        Me.Programa_Rótulo.Visible = True
        Me.Programa.Visible = True
        Me.Repasse_Rótulo.Visible = True
        Me.Repasse.Visible = True
    Else
        Me.Programa_Rótulo.Visible = False
        Me.Programa.Visible = False
        Me.Repasse_Rótulo.Visible = False
        Me.Repasse.Visible = False
    End If
    Me.BP.SetFocus
End Sub

Private Sub Parcial_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizar a ordem de tabulação e adicionar funções específicas às teclas de navegação entre controles

    Select Case KeyCode
        Case Is = 9
            If Shift <> 1 Then
                If Me.Repasse.Visible = True And Me.Programa.Visible = True Then
                    KeyCode = 0
                    Me.Repasse.SetFocus
                    GoTo Finish
                Else
                    KeyCode = 0
                    Me.BP.SetFocus
                End If
            Else
                KeyCode = 0
                Me.Pirâmide.SetFocus
            End If
        Case Is = 13, 32
            KeyCode = 0
            Me.Parcial.Value = True
                If Me.Repasse.Visible = True Then
                    Me.Repasse.SetFocus
                Else
                    Me.BP.SetFocus
                End If
        Case Is = 8, 46
            KeyCode = 0
            Me.Parcial.Value = False
                If Me.Repasse.Visible = True Then
                    Me.Repasse.SetFocus
                Else
                    Me.BP.SetFocus
                End If
        Case Is = 39, 40
                If Me.Repasse.Visible = True And Me.Programa.Visible = True Then
                    KeyCode = 0
                    Me.Repasse.SetFocus
                    GoTo Finish
                Else
                    KeyCode = 0
                    Me.BP.SetFocus
                End If
        Case Else
            Exit Sub
    End Select
Finish:
End Sub

Private Sub Pirâmide_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para adicionar funções específicas para as teclas
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub Produzido_Exit(Cancel As Integer)
'macro para verificar o correto lançamento dos dados nos campos "produzido" e "refugado"

    If Me!Produzido < Me!Refugado Then
        MsgBox "A quantidade produzida não pode ser menor que a quantidade refugada!", vbOKOnly, "WIP MFPA"
        Me.Produzido.Value = PrevProdVal
        Me.Produzido.SetFocus
        PrevProdVal = 0
    End If
End Sub

Private Sub Produzido_GotFocus()
'macro para verificar o correto lançamento dos dados nos campos "produzido" e "refugado"
    PrevProdVal = Nz(Me.Produzido.Value, 0)
End Sub

Private Sub Produzido_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para adicionar funções específicas para as teclas
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub Programa_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizar a ordem de tabulação e adicionar funções específicas às teclas de navegação entre controles

    Select Case KeyCode
        Case Is = 13, 39, 40
            KeyCode = 0
            Me.BP.SetFocus
        Case Is = 37, 38
            KeyCode = 0
            Me.Repasse.SetFocus
        Case Is = 9
            If Shift <> 1 Then
                KeyCode = 0
                Me.BP.SetFocus
            Else
                KeyCode = 0
                Me.Repasse.SetFocus
            End If
        Case Is = 27
            Desfazer_Click
            Me.BP.SetFocus
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Refugado_GotFocus()
'macro para verificar o correto lançamento dos dados nos campos "produzido" e "refugado"
    PrevRefVal = Nz(Me.Refugado.Value, 0)
End Sub

Private Sub Refugado_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para adicionar funções específicas para as teclas
    If KeyCode = 27 Then
      Desfazer_Click
    End If
End Sub

Private Sub Refugado_LostFocus()
'macro para verificar o correto lançamento dos dados nos campos "produzido" e "refugado"

    If Me!Produzido < Me!Refugado Then
        MsgBox "A quantidade refugada não pode ser maior que a quantidade produzida!", vbOKOnly, "WIF MFPA"
        Me.Refugado.Value = PrevRefVal
        Me.Produzido.SetFocus
        PrevRefVal = 0
    End If
End Sub

Private Sub Repasse_KeyDown(KeyCode As Integer, Shift As Integer)
'macro para sistematizar a ordem de tabulação e adicionar funções específicas às teclas de navegação entre controles

Select Case KeyCode
    Case Is = 13
    Me.Repasse.Value = True
    Case Is = 8
    Me.Repasse.Value = False
    Case Else
End Select
End Sub

Private Sub SAIR_Click()
'macro que salva e fecha o arquivo

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

    Me.Ícone_Sair.SetFocus
    If Me.NewRecord = True And Me.Dirty = False Then
    SelectOPA = True
    End If
    If Not Me.NewRecord And Me.Dirty = True Then
        If MsgBox("Você realmente deseja sair? As alterações serão descartadas.", vbYesNo, "WIP MFPA") = vbYes Then
            SelectOPA = True
            Me.Undo
            EventExec = True
            DoCmd.Quit acQuitSaveAll
        Else
            Exit Sub
        End If
    End If
    If Me.Dirty = True Or SelectOPA = False Then
        If MsgBox("Você realmente deseja sair do aplicativo?", vbYesNo, "WIP MFPA") = vbYes Then
            SelectOPA = True
            DoCmd.Quit acQuitSaveNone
        End If
    End If
    If MsgBox("Você realmente deseja sair do aplicativo?", vbYesNo, "WIP MFPA") = vbYes Then
        SelectOPA = True
        EventExec = True
        DoCmd.Quit acQuitSaveAll
    End If
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

Private Sub SALVAR_Click()
'macro para salvar o registro

Dim Operação As String
Dim StrSQL As String
Dim db As DAO.Database
Dim Historico As Recordset
Dim Encerrado As Recordset
Dim FluxoBD As Recordset, rst2 As Recordset, rst3 As Recordset, rst4 As Recordset, rst5 As Recordset
Dim FLUXO As String, TECNOLOGIA As String, Agrupamento As String
Dim IDFluxo As Integer
Dim Next_Oper As String
Dim Curr_Oper As String, Curr_AN As String, Curr_OPA As Long
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String

On Error GoTo ErrHandler

Set db = CurrentDb
Set Historico = db.OpenRecordset("Histórico")
Set Encerrado = db.OpenRecordset("Encerrado")

Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
LocalControl = Me.ActiveControl.Name

If Me.NewRecord = False Then
    Operação = PrevOperVal
Else
    Curr_AN = Nz(Me.AN.Value)
    Curr_OPA = Nz(Me.OPA.Value)
    Curr_Oper = Nz(Me.Operação.Value)
    Operação = Nz(Me.Operação.Value)
    New_Rec = True
End If

EventExec = True
EventExec2 = True
Me.Ícone_Salvar.SetFocus

    If Me.Produzido.Value - Me.Refugado.Value > 0 Then
        Select Case Operação
            Case Is = "PRETEJAMENTO"
                If IsNull(Me.OPA.Value) Or IsNull(Me.AN.Value) Or IsNull(Me.Produzido.Value) Or _
                IsNull(Me.Refugado.Value) Or IsNull(Me.Total.Value) Or IsNull(Me.Pirâmide.Value) Or _
                IsNull(Me.Máquina.Value) Or IsNull(Me.BP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
                IsNull(Me.Nome_do_Colaborador.Value) Or IsNull(Me.NP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
                IsNull(Me.Data.Value) Or IsNull(Me.Programa.Value) Or IsNull(Me.Repasse.Value) Then
                    MsgBox "Por favor, preencha todos os campos antes de salvar!", vbOKOnly, "WIP MFPA"
                    EventExec = False
                    EventExec2 = False
                    Exit Sub
                Else
                    If Me.Dirty = True Or SelectOPA = False Then
                        If Me.NewRecord = False Then
                            If Me.Operação.Value <> PrevOperVal Then
                                Me.Programa.Value = ""
                                Me.Repasse.Value = 0
                                If Me.OPA.Value <> PrevOPAVal Then
                                    StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Programa = '" & Me.Programa & "', Histórico.Repasse = '" & Me.Repasse & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    Me.Dirty = False
                                    StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                            ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                            ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Fluxo.OPA = " & PrevOPAVal
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                Else
                                    StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Programa = '" & Me.Programa & "', Histórico.Repasse = '" & Me.Repasse & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    Me.Dirty = False
                                    StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                            ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                            ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Fluxo.OPA = " & Me.OPA
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                End If
                            Else
                               If Me.OPA.Value <> PrevOPAVal Then
                                    StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.Programa = '" & Me.Programa & "', Histórico.Repasse = '" & Me.Repasse & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    Me.Dirty = False
                                    StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                            ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                            ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Fluxo.OPA = " & PrevOPAVal
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                Else
                                    StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.Programa = '" & Me.Programa & "', Histórico.Repasse = '" & Me.Repasse & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA =" & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    Me.Dirty = False
                                    StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                            ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                            ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Fluxo.OPA = " & Me.OPA
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                End If
                            End If
                        Else
                            Historico.AddNew
                            Historico!OPA = Me.OPA
                            Historico!AN = Me.AN
                            Historico!Colaborador = Me.Nome_do_Colaborador
                            Historico!NP = Me.NP
                            Historico!Data = Me.Data
                            Historico!Hora = Me.Hora
                            Historico!Operação = Me.Operação
                            Historico!Máquina = Me.Máquina
                            Historico!BP = Me.BP
                            Historico!Produzido = Me.Produzido
                            Historico!Refugado = Me.Refugado
                            Historico!Total = Me.Total
                            Historico!Pirâmide = Me.Pirâmide
                            Historico!Parcial = Me.Parcial
                            Historico!CC = Me.Centro_de_Custo
                            Historico!Programa = Me.Programa
                            Historico!Repasse = Me.Repasse
                            Historico.Update
                            StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            Me.Dirty = False
                            StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                    ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                    ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                    ", Fluxo.Data = '" & Me.Data & "', Fluxo.Hora = '" & Me.Hora & "'" & _
                                    " WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            DoCmd.Save
                            MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                            DoCmd.RunCommand acCmdRecordsGoToNew
                            Me.BP.SetFocus
                        End If
                    Else
                        MsgBox "Não foi realizada nenhuma alteração no registro!", vbOKOnly, "WIP MFPA"
                        EventExec = False
                        EventExec2 = False
                        Exit Sub
                    End If
                End If
            Case Is = "FORMADORA", "ENROLAMENTO", "ESTAMPAGEM", "EXPANSOR"
                If IsNull(Me.OPA.Value) Or IsNull(Me.AN.Value) Or IsNull(Me.Produzido.Value) Or _
                IsNull(Me.Refugado.Value) Or IsNull(Me.Total.Value) Or IsNull(Me.Pirâmide.Value) Or _
                IsNull(Me.Máquina.Value) Or IsNull(Me.BP.Value) Or _
                IsNull(Me.Nome_do_Colaborador.Value) Or IsNull(Me.NP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
                IsNull(Me.Data.Value) Then
                    MsgBox "Por favor, preencha todos os campos antes de salvar!", vbOKOnly, "WIP MFPA"
                    EventExec = False
                    EventExec2 = False
                    Exit Sub
                Else
                    If Me.Dirty = True Or SelectOPA = False Then
                        If Me.NewRecord = False Then
                            If Me.OPA.Value <> PrevOPAVal Then
                                StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                Me.Dirty = False
                                StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                        ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                        ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Fluxo.OPA = " & PrevOPAVal
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            Else
                                Me.Dirty = False
                                StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                        ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                        ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Fluxo.OPA = " & PrevOPAVal
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            End If
                        Else
                            Historico.AddNew
                            Historico!OPA = Me.OPA
                            Historico!AN = Me.AN
                            Historico!Colaborador = Me.Nome_do_Colaborador
                            Historico!NP = Me.NP
                            Historico!Data = Me.Data
                            Historico!Hora = Me.Hora
                            Historico!Operação = Me.Operação
                            Historico!Máquina = Me.Máquina
                            Historico!BP = Me.BP
                            Historico!Produzido = Me.Produzido
                            Historico!Refugado = Me.Refugado
                            Historico!Total = Me.Total
                            Historico!Pirâmide = Me.Pirâmide
                            Historico!Parcial = Me.Parcial
                            Historico!CC = Me.Centro_de_Custo
                            Historico.Update
                            Me.Dirty = False
                            StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                    ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                    ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                    ", Fluxo.Data = '" & Me.Data & "', Fluxo.Hora = '" & Me.Hora & "'" & _
                                    " WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            DoCmd.Save
                            MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                            DoCmd.RunCommand acCmdRecordsGoToNew
                            Me.BP.SetFocus
                        End If
                    Else
                        MsgBox "Não foi realizada nenhuma alteração no registro!", vbOKOnly, "WIP MFPA"
                        EventExec = False
                        EventExec2 = False
                        Exit Sub
                    End If
                End If
            Case Is = "ESTOQUE"
                If IsNull(Me.OPA.Value) Or IsNull(Me.AN.Value) Or IsNull(Me.Produzido.Value) Or _
                IsNull(Me.Refugado.Value) Or IsNull(Me.Total.Value) Or IsNull(Me.Pirâmide.Value) Or _
                IsNull(Me.Máquina.Value) Or IsNull(Me.BP.Value) Or IsNull(Me.Nome_do_Colaborador.Value) Or _
                IsNull(Me.NP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
                IsNull(Me.Data.Value) Then
                    MsgBox "Por favor, preencha todos os campos antes de salvar!", vbOKOnly, "WIP MFPA"
                    EventExec = False
                    EventExec2 = False
                    Exit Sub
                Else
                    If MsgBox("ATENÇÃO! Após o salvamento do registro nessa operação NÃO será mais possível realizar alterações! Você " & _
                    "realmente deseja continuar?", vbYesNo, "WIP MFPA") = vbYes Then
                        If Me.Dirty = True Or SelectOPA = False Then
                            If Me.NewRecord = False Then
                                If Me.OPA.Value <> PrevOPAVal Then
                                    StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    StrSQL = "UPDATE Encerrado SET Encerrado.OPA = " & Me.OPA & ", Encerrado.AN = '" & Me.AN & "'" & _
                                            ", Encerrado.Colaborador = '" & Me.Nome_do_Colaborador & "', Encerrado.NP = '" & Me.NP & "'" & _
                                            ", Encerrado.Produzido = '" & Me.Produzido & "', Encerrado.Refugado = '" & Me.Refugado & "'" & _
                                            ", Encerrado.Total = '" & Me.Total & "', Encerrado.Operação = '" & Me.Operação & "'" & _
                                            ", Encerrado.BP = '" & Me.BP & "', Encerrado.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Encerrado.Parcial = '" & Me.Parcial & "'" & _
                                            ", Encerrado.Máquina = '" & Me.Máquina & "', Encerrado.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Encerrado.OPA = " & PrevOPAVal & " AND Encerrado.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    Me.Dirty = False
                                    StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                Else
                                    Me.Dirty = False
                                    StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                            ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                            ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                            ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                            ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                            ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    StrSQL = "UPDATE Encerrado SET Encerrado.OPA = " & Me.OPA & ", Encerrado.AN = '" & Me.AN & "'" & _
                                            ", Encerrado.Colaborador = '" & Me.Nome_do_Colaborador & "', Encerrado.NP = '" & Me.NP & "'" & _
                                            ", Encerrado.Produzido = '" & Me.Produzido & "', Encerrado.Refugado = '" & Me.Refugado & "'" & _
                                            ", Encerrado.Total = '" & Me.Total & "', Encerrado.Operação = '" & Me.Operação & "'" & _
                                            ", Encerrado.BP = '" & Me.BP & "', Encerrado.Pirâmide = '" & Me.Pirâmide & "'" & _
                                            ", Encerrado.Parcial = '" & Me.Parcial & "'" & _
                                            ", Encerrado.Máquina = '" & Me.Máquina & "', Encerrado.CC = '" & Me.Centro_de_Custo & "'" & _
                                            " WHERE Encerrado.OPA = " & PrevOPAVal & " AND Encerrado.Operação = '" & PrevOperVal & "'"
                                    DoCmd.RunSQL StrSQL
                                    StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                                    DoCmd.RunSQL StrSQL
                                    DoCmd.Save
                                    MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                    DoCmd.RunCommand acCmdRecordsGoToNew
                                    Me.BP.SetFocus
                                End If
                            Else
                                Historico.AddNew
                                Historico!OPA = Me.OPA
                                Historico!AN = Me.AN
                                Historico!Colaborador = Me.Nome_do_Colaborador
                                Historico!NP = Me.NP
                                Historico!Data = Me.Data
                                Historico!Hora = Me.Hora
                                Historico!Operação = Me.Operação
                                Historico!Máquina = Me.Máquina
                                Historico!BP = Me.BP
                                Historico!Produzido = Me.Produzido
                                Historico!Refugado = Me.Refugado
                                Historico!Total = Me.Total
                                Historico!Pirâmide = Me.Pirâmide
                                Historico!Parcial = Me.Parcial
                                Historico!CC = Me.Centro_de_Custo
                                Historico.Update
                                Encerrado.AddNew
                                Encerrado!OPA = Me.OPA
                                Encerrado!AN = Me.AN
                                Encerrado!Colaborador = Me.Nome_do_Colaborador
                                Encerrado!NP = Me.NP
                                Encerrado!Data = Me.Data
                                Encerrado!Hora = Me.Hora
                                Encerrado!Operação = Me.Operação
                                Encerrado!Máquina = Me.Máquina
                                Encerrado!BP = Me.BP
                                Encerrado!Produzido = Me.Produzido
                                Encerrado!Refugado = Me.Refugado
                                Encerrado!Total = Me.Total
                                Encerrado!Pirâmide = Me.Pirâmide
                                Encerrado!Parcial = Me.Parcial
                                Encerrado!CC = Me.Centro_de_Custo
                                Encerrado.Update
                                StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                                DoCmd.RunSQL StrSQL
                                Me.Undo
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                Me.BP.SetFocus
                            End If
                        Else
                            MsgBox "Não foi realizada nenhuma alteração no registro!", vbOKOnly, "WIP MFPA"
                            EventExec = False
                            EventExec2 = False
                            Exit Sub
                        End If
                    Else
                        EventExec = False
                        EventExec2 = False
                        Exit Sub
                    End If
                End If
            Case Else
                If IsNull(Me.OPA.Value) Or IsNull(Me.AN.Value) Or IsNull(Me.Produzido.Value) Or _
                IsNull(Me.Refugado.Value) Or IsNull(Me.Total.Value) Or IsNull(Me.Pirâmide.Value) Or _
                IsNull(Me.Máquina.Value) Or IsNull(Me.BP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
                IsNull(Me.Nome_do_Colaborador.Value) Or IsNull(Me.NP.Value) Or IsNull(Me.Hora.Value) Or _
                IsNull(Me.Data.Value) Then
                    MsgBox "Por favor, preencha todos os campos antes de salvar!", vbOKOnly, "WIP MFPA"
                    EventExec = False
                    EventExec2 = False
                    Exit Sub
                Else
                    If Me.Dirty = True Or SelectOPA = False Then
                        If Me.NewRecord = False Then
                            If Me.OPA.Value <> PrevOPAVal Then
                                StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                Me.Dirty = False
                                StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                        ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                        ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Fluxo.OPA = " & PrevOPAVal
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            Else
                                Me.Dirty = False
                                StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                        ", Fluxo.Total = '" & Me.Total & "', Fluxo.Operação = '" & Me.Operação & "'" & _
                                        ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Fluxo.OPA = " & PrevOPAVal
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            End If
                        Else
                            Historico.AddNew
                            Historico!OPA = Me.OPA
                            Historico!AN = Me.AN
                            Historico!Colaborador = Me.Nome_do_Colaborador
                            Historico!NP = Me.NP
                            Historico!Data = Me.Data
                            Historico!Hora = Me.Hora
                            Historico!Operação = Me.Operação
                            Historico!Máquina = Me.Máquina
                            Historico!BP = Me.BP
                            Historico!Produzido = Me.Produzido
                            Historico!Refugado = Me.Refugado
                            Historico!Total = Me.Total
                            Historico!Pirâmide = Me.Pirâmide
                            Historico!Parcial = Me.Parcial
                            Historico!CC = Me.Centro_de_Custo
                            Historico.Update
                            StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            Me.Refresh
                            Me.Dirty = False
                            StrSQL = "UPDATE Fluxo SET Fluxo.Colaborador = '" & Me.Nome_do_Colaborador & "'" & _
                                    ", Fluxo.Total = " & Me.Total & ", Fluxo.Operação = '" & Me.Operação & "'" & _
                                    ", Fluxo.Máquina = '" & Me.Máquina & "', Fluxo.CC = '" & Me.Centro_de_Custo & "'" & _
                                    ", Fluxo.Data = '" & Me.Data & "', Fluxo.Hora = '" & Me.Hora & "'" & _
                                    " WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            DoCmd.Save
                            MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                            DoCmd.RunCommand acCmdRecordsGoToNew
                            Me.BP.SetFocus
                        End If
                    Else
                        MsgBox "Não foi realizada nenhuma alteração no registro!", vbOKOnly, "WIP MFPA"
                        EventExec = False
                        EventExec2 = False
                        Exit Sub
                    End If
                End If
        End Select
    Else
        If IsNull(Me.OPA.Value) Or IsNull(Me.AN.Value) Or IsNull(Me.Produzido.Value) Or _
    IsNull(Me.Refugado.Value) Or IsNull(Me.Total.Value) Or IsNull(Me.Pirâmide.Value) Or _
    IsNull(Me.Máquina.Value) Or IsNull(Me.BP.Value) Or IsNull(Me.Centro_de_Custo.Value) Or _
    IsNull(Me.Nome_do_Colaborador.Value) Or IsNull(Me.NP.Value) Or IsNull(Me.Data.Value) Then
            MsgBox "Por favor, preencha todos os campos antes de salvar!", vbOKOnly, "WIP MFPA"
            EventExec = False
            EventExec2 = False
            Exit Sub
        Else
            If MsgBox("A quantidade refugada representa 100% do lote. Isso significa que a ordem será encerrada. " & _
            "Deseja continuar?", vbYesNo, "WIP MFPA") = vbYes Then
                If MsgBox("ATENÇÃO! Após o salvamento do registro nessa operação NÃO será mais possível realizar alterações! Você " & _
                "realmente deseja continuar?", vbYesNo, "WIP MFPA") = vbYes Then
                    If Me.Dirty = True Or SelectOPA = False Then
                        If Me.NewRecord = False Then
                            If Me.OPA.Value <> PrevOPAVal Then
                                StrSQL = "UPDATE Histórico SET Histórico.OPA = " & Me.OPA & ", Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & PrevOPAVal & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                StrSQL = "UPDATE Encerrado SET Encerrado.OPA = " & Me.OPA & ", Encerrado.AN = '" & Me.AN & "'" & _
                                        ", Encerrado.Colaborador = '" & Me.Nome_do_Colaborador & "', Encerrado.NP = '" & Me.NP & "'" & _
                                        ", Encerrado.Produzido = '" & Me.Produzido & "', Encerrado.Refugado = '" & Me.Refugado & "'" & _
                                        ", Encerrado.Total = '" & Me.Total & "', Encerrado.Operação = '" & Me.Operação & "'" & _
                                        ", Encerrado.BP = '" & Me.BP & "', Encerrado.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Encerrado.Parcial = '" & Me.Parcial & "'" & _
                                        ", Encerrado.Máquina = '" & Me.Máquina & "', Encerrado.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Encerrado.OPA = " & PrevOPAVal & " AND Encerrado.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                Me.Dirty = False
                                StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            Else
                                Me.Dirty = False
                                StrSQL = "UPDATE Histórico SET Histórico.AN = '" & Me.AN & "'" & _
                                        ", Histórico.Colaborador = '" & Me.Nome_do_Colaborador & "', Histórico.NP = '" & Me.NP & "'" & _
                                        ", Histórico.Produzido = '" & Me.Produzido & "', Histórico.Refugado = '" & Me.Refugado & "'" & _
                                        ", Histórico.Total = '" & Me.Total & "', Histórico.Operação = '" & Me.Operação & "'" & _
                                        ", Histórico.BP = '" & Me.BP & "', Histórico.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Histórico.Parcial = '" & Me.Parcial & "'" & _
                                        ", Histórico.Máquina = '" & Me.Máquina & "', Histórico.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Histórico.OPA = " & Me.OPA & " AND Histórico.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                StrSQL = "UPDATE Encerrado SET Encerrado.OPA = " & Me.OPA & ", Encerrado.AN = '" & Me.AN & "'" & _
                                        ", Encerrado.Colaborador = '" & Me.Nome_do_Colaborador & "', Encerrado.NP = '" & Me.NP & "'" & _
                                        ", Encerrado.Produzido = '" & Me.Produzido & "', Encerrado.Refugado = '" & Me.Refugado & "'" & _
                                        ", Encerrado.Total = '" & Me.Total & "', Encerrado.Operação = '" & Me.Operação & "'" & _
                                        ", Encerrado.BP = '" & Me.BP & "', Encerrado.Pirâmide = '" & Me.Pirâmide & "'" & _
                                        ", Encerrado.Parcial = '" & Me.Parcial & "'" & _
                                        ", Encerrado.Máquina = '" & Me.Máquina & "', Encerrado.CC = '" & Me.Centro_de_Custo & "'" & _
                                        " WHERE Encerrado.OPA = " & PrevOPAVal & " AND Encerrado.Operação = '" & PrevOperVal & "'"
                                DoCmd.RunSQL StrSQL
                                StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                                DoCmd.RunSQL StrSQL
                                DoCmd.Save
                                MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                                DoCmd.RunCommand acCmdRecordsGoToNew
                                Me.BP.SetFocus
                            End If
                        Else
                            Historico.AddNew
                            Historico!OPA = Me.OPA
                            Historico!AN = Me.AN
                            Historico!Colaborador = Me.Nome_do_Colaborador
                            Historico!NP = Me.NP
                            Historico!Data = Me.Data
                            Historico!Hora = Me.Hora
                            Historico!Operação = Me.Operação
                            Historico!Máquina = Me.Máquina
                            Historico!BP = Me.BP
                            Historico!Produzido = Me.Produzido
                            Historico!Refugado = Me.Refugado
                            Historico!Total = Me.Total
                            Historico!Pirâmide = Me.Pirâmide
                            Historico!Parcial = Me.Parcial
                            Historico!CC = Me.Centro_de_Custo
                            Historico.Update
                            Encerrado.AddNew
                            Encerrado!OPA = Me.OPA
                            Encerrado!AN = Me.AN
                            Encerrado!Colaborador = Me.Nome_do_Colaborador
                            Encerrado!NP = Me.NP
                            Encerrado!Data = Me.Data
                            Encerrado!Hora = Me.Hora
                            Encerrado!Operação = Me.Operação
                            Encerrado!Máquina = Me.Máquina
                            Encerrado!BP = Me.BP
                            Encerrado!Produzido = Me.Produzido
                            Encerrado!Refugado = Me.Refugado
                            Encerrado!Total = Me.Total
                            Encerrado!Pirâmide = Me.Pirâmide
                            Encerrado!Parcial = Me.Parcial
                            Encerrado!CC = Me.Centro_de_Custo
                            Encerrado.Update
                            StrSQL = "DELETE Fluxo.*, Fluxo.OPA FROM Fluxo WHERE Fluxo.OPA = " & Me.OPA
                            DoCmd.RunSQL StrSQL
                            Me.Undo
                            DoCmd.Save
                            MsgBox "Registro salvo com sucesso!", vbOKOnly, "WIP MFPA"
                            Me.BP.SetFocus
                        End If
                    Else
                        MsgBox "Não foi realizada nenhuma alteração no registro!", vbOKOnly, "WIP MFPA"
                        EventExec = False
                        EventExec2 = False
                        Exit Sub
                    End If
                Else
                    EventExec = False
                    EventExec2 = False
                    Exit Sub
                End If
            End If
        End If
    End If
If Curr_Oper = "ESTOQUE" Then GoTo Final
If New_Rec = True Then
    Set FluxoBD = db.OpenRecordset("Fluxo")
    Set rst2 = db.OpenRecordset("Fluxo por Tecnologia")
    Set rst3 = db.OpenRecordset("Sequência de Operações por Fluxo e Tecnologia")
    Set rst4 = db.OpenRecordset("BD Fluxo por Tecnologia")
    Set rst5 = db.OpenRecordset("BD Agrupamentos")
    rst4.MoveFirst
    Do Until rst4!AN = Curr_AN
        rst4.MoveNext
        If rst4.EOF Then
            MsgBox "Avisar programação que a ficha " & Curr_AN & " não está cadastrada no sistema!", vbCritical, "WIP MFPA"
            Exit Do
        End If
    Loop
    If rst4.EOF Then GoTo Final
    FLUXO = rst4!FLUXO
    TECNOLOGIA = rst4!TECNOLOGIA
    rst2.MoveFirst
    Do Until rst2!TECNOLOGIA = TECNOLOGIA And rst2!FLUXO = FLUXO
        rst2.MoveNext
        If rst2.EOF Then
            Exit Do
        End If
    Loop
    IDFluxo = rst2!ID_Fluxo
    rst3.MoveFirst
    Do Until rst3!ID_Fluxo = IDFluxo And rst3!TECNOLOGIA = TECNOLOGIA And rst3!FLUXO = FLUXO And rst3!Operação = Curr_Oper
    rst3.MoveNext
        If rst3.EOF Then
            Exit Do
        End If
    Loop
    If rst3.EOF Then
        MsgBox "Avisar programação que a ficha " & Curr_AN & " não está cadastrada no sistema!", vbCritical, "WIP MFPA"
        GoTo Final
    End If
    rst3.MoveNext
    Next_Oper = rst3!Operação
    Do Until rst5!AN = Curr_AN
        rst5.MoveNext
        If rst5.EOF Then
            Exit Do
        End If
    Loop
    Agrupamento = rst5!Agrupamento
    FluxoBD.MoveFirst
    Do Until FluxoBD!AN = Curr_AN And FluxoBD!OPA = Curr_OPA
        FluxoBD.MoveNext
        If FluxoBD.EOF Then
            Exit Do
        End If
    Loop
    FluxoBD.Edit
    FluxoBD!Próxima_Operação = Next_Oper
    FluxoBD!Agrupamento = Agrupamento
    FluxoBD.Update
    Set FluxoBD = Nothing
    Set rst2 = Nothing
    Set rst3 = Nothing
    Set rst4 = Nothing
    Set rst5 = Nothing
    Set db = Nothing
End If
Final:
Set Historico = Nothing
Set Encerrado = Nothing
Me.Refresh
SelectOPA = True
EventExec = False
EventExec2 = False
New_Rec = False
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

Private Sub Total_AfterUpdate()
'macro para impedir o lançamento de dados incorretos
If Me.Total.Value < 0 Then
    Me.Produzido.Value = 0
    Me.Refugado.Value = 0
End If
End Sub

Private Sub Total_Enter()
'macro para sistematizar a ordem de tabulação com a ordem correta de lançamento de dados no apontamento
    Me.Pirâmide.SetFocus
End Sub

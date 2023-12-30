VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Painel_do_Gestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Atualizar_Painel_Click()
'macro para atualizar os dados do painel
    Me.Atualizar_Painel_Ícone.SetFocus
    EventExec2 = True
    Me.Requery
    Me.Refresh
    EventExec2 = False
End Sub

Private Sub Atualizar_Painel_Ícone_Click()
'macro para atualizar os dados do painel
Atualizar_Painel_Click
End Sub

Private Sub Exportar_Excel_Btn_Click()
'macro para exportar o arquivo de relatório de produção em formato excel

Dim StrSQL As String
Dim InitialDate As Variant
Dim EndDate As Variant

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
    
    Me.Exportar_Excel_Ícone.SetFocus
    Do
        InitialDate = InputBox("Insira a data de início: ", "WIP MFPA")
        If Nz(InitialDate, "") = "" Then
            Exit Sub
        End If
        InitialDate = Format(InitialDate, "mm/dd/yyyy")
        If IsDate(InitialDate) Then Exit Do
    Loop
    Do
        EndDate = InputBox("Insira a data final: ", "WIP MFPA")
        If Nz(EndDate, "") = "" Then
            Exit Sub
        End If
        EndDate = Format(EndDate, "mm/dd/yyyy")
        If EndDate < InitialDate Then
            MsgBox "A data final deve ser maior que a data inicial!", vbOKOnly, "WIP MFPA"
        End If
        If IsDate(EndDate) And EndDate >= InitialDate Then Exit Do
    Loop
    DoCmd.Hourglass True
    StrSQL = "DELETE Apontamento.* FROM Apontamento;"
    DoCmd.RunSQL StrSQL
    StrSQL = "INSERT INTO Apontamento ( Data, Hora, Operação, Máquina, Produzido, Refugado, Refugo ) " & _
            "SELECT Histórico.Data, Histórico.Hora, Histórico.Operação, Histórico.Máquina, Histórico.Produzido, Histórico.Refugado, FormatPercent([Refugado]/[Produzido],2) AS Refugo " & _
            "FROM Histórico WHERE (((Histórico.Data) Between #" & InitialDate & "# And #" & EndDate & "#)) " & _
            "GROUP BY Histórico.Data, Histórico.Hora, Histórico.Operação, Histórico.Máquina, Histórico.Produzido, Histórico.Refugado, FormatPercent([Refugado]/[Produzido],2);"
    DoCmd.RunSQL StrSQL
    DoCmd.Hourglass False
    MsgBox "Relatório de Produção exportado com sucesso!", vbOKOnly, "WIP MFPA"
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

Private Sub Exportar_Excel_Ícone_Click()
'macro para exportar o arquivo de relatório de produção em formato excel
Exportar_Excel_Btn_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para carregar a versão do programa

Dim db As DAO.Database
Dim rst As Recordset
Dim Version As String

Set db = CurrentDb
Set rst = db.OpenRecordset("StationVersion")

rst.MoveFirst
Version = rst!VersionName
Me.Versao_Sistema = Version

Forms!Painel_do_Gestor.AllowDeletions = False

Set db = Nothing
Set rst = Nothing

End Sub

Private Sub Ícone_Sair_Click()
'macro para sair do app
SAIR_Click
End Sub

Private Sub Quadro_BD_Click()
'macro para sistematizar a alternânica do modo de visualização do banco de dados entre formulário e tabelas
    If Me.Quadro_BD.Value = 1 Then
        Me.Guia_Tabelas.Visible = False
        Me.Guia_Form.Visible = True
    Else
        Me.Guia_Tabelas.Visible = True
        Me.Guia_Form.Visible = False
    End If
End Sub

Private Sub Report_Detalhado_Botao_Click()
'macro para chamar a subrotina de criação do report wip detalhado e exibí-lo em modo de visualização de impressão

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String
Dim Attpath As String

On Error GoTo ErrHandler
Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
LocalControl = Me.ActiveControl.Name

    Me.Visible = False
    Call WIPRpt
    Attpath = CurrentProject.Path & "\Relatório Detalhado WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Relatório Detalhado WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    DoCmd.OpenReport "Relatório Detalhado WIP", acViewPreview
    DoCmd.RunCommand acCmdZoom75
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

Private Sub Report_Erro_Btn_Click()
'macro para exportar em pdf o relatório de erros e exibí-lo em modo de visualização de impressão

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String
Dim Attpath As String
Dim StrSQL As String

On Error GoTo ErrHandler
Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
LocalControl = Me.ActiveControl.Name
    
    Me.Report_Erro_Ícone.SetFocus
    If MsgBox("Deseja atualizar e exportar novo relatório de erros do sistema?", vbYesNo, "WIP MFPA") = vbYes Then
        If Backlog.RecordCount = 0 Then
            MsgBox "Não há erros para serem exibidos!", vbOKOnly, "WIP MFPA"
        Else
            Me.Visible = False
            Attpath = CurrentProject.Path & "\Relatório de Erros.pdf"
            DoCmd.OutputTo acOutputReport, "Relatório de Erros", acFormatPDF, Attpath, False, , , acExportQualityPrint
            DoCmd.OpenReport "Relatório de Erros", acViewPreview
            DoCmd.RunCommand acCmdZoom75
            StrSQL = "DELETE ErrorBacklog.* FROM ErrorBacklog;"
            DoCmd.RunSQL StrSQL
        End If
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

Private Sub Report_Erro_Ícone_Click()
'macro para exportar em pdf o relatório de erros e exibí-lo em modo de visualização de impressão
Report_Erro_Btn_Click
End Sub

Private Sub Report_WIP_Botao_Click()
'macro para chamar a subrotina de criação do report wip e exibí-lo em modo de visualização de impressão

Dim db As DAO.Database
Dim Backlog As Recordset
Dim FileName As String
Dim LocalObject As String
Dim LocalControl As String
Dim Attpath As String

On Error GoTo ErrHandler
Set db = CurrentDb
Set Backlog = db.OpenRecordset("ErrorBacklog")
FileName = CurrentProject.Name
LocalObject = CurrentObjectName
LocalControl = Me.ActiveControl.Name

    Me.Visible = False
    Call WIPRpt
    Attpath = CurrentProject.Path & "\Report WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Report WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    DoCmd.OpenReport "Report WIP", acViewPreview
    DoCmd.RunCommand acCmdZoom75
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

Private Sub Report_WIP_Detalhado_Outlook_Click()
'macro para salvar o relatório em pdf e enviar por email de acordo com a lista de email disponibilizada

    Dim db As DAO.Database
    Dim rst As Recordset
    Dim EmailSubj As String
    Dim EmailBody As String
    Dim Destination As String
    Dim olApp As Outlook.Application
    Dim olMsg As Outlook.MailItem
    Dim MyAtt As Object
    Dim Attpath As String
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
    
    Call WIPRpt
    Attpath = CurrentProject.Path & "\Relatório Detalhado WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Relatório Detalhado WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    Set db = CurrentDb
    Set rst = db.OpenRecordset("EmailContent")
    If rst.RecordCount = 0 Then
        MsgBox "Favor registrar padrão de corpo do email!", vbCritical, "WIP MFPA"
        rst.Close
        Exit Sub
    End If
    rst.MoveFirst
    Do Until rst!ID = 2
        rst.MoveNext
    Loop
    EmailSubj = rst!Assunto
    EmailBody = rst!Corpo
    rst.Close
    Set rst = db.OpenRecordset("EmailList")
    If rst.RecordCount = 0 Then
        MsgBox "Favor cadastrar endereços de email de destinatário!", vbCritical, "WIP MFPA"
        rst.Close
        Exit Sub
    End If
    rst.MoveFirst
    Destination = rst!Email
    Set olApp = CreateObject("Outlook.Application")
    Set olMsg = olApp.CreateItem(olMailItem)
    Do
        rst.MoveNext
        If rst.EOF Then Exit Do
        Destination = Destination & " ; " & rst!Email
    Loop
    olMsg.To = "" & Destination
    olMsg.Subject = EmailSubj
    olMsg.Body = EmailBody
    Set MyAtt = olMsg.Attachments
    MyAtt.Add "" & Attpath
    olMsg.Send
    Set olMsg = Nothing
    Set rst = Nothing
    Set db = Nothing
    MsgBox "Report enviado via Email com sucesso!", vbOKOnly, "WIP MFPA"
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

Private Sub Report_WIP_Detalhado_Print_Botao_Click()
'macro para chamar a subrotina de criação do report wip detalhado e imprimí-lo

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

    Me.Visible = False
    Call WIPRpt
    DoCmd.OpenReport "Relatório Detalhado WIP", acViewPreview
    DoCmd.RunCommand acCmdPrint
    DoCmd.Close acReport, "Relatório Detalhado WIP"
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Is = 2501
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

Private Sub Report_WIP_Outlook_Click()
'macro para salvar o relatório em pdf e enviar por email de acordo com a lista de email disponibilizada

    Dim db As DAO.Database
    Dim rst As Recordset
    Dim EmailSubj As String
    Dim EmailBody As String
    Dim Destination As String
    Dim olApp As Outlook.Application
    Dim olMsg As Outlook.MailItem
    Dim MyAtt As Object
    Dim Attpath As String
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
    
    Call WIPRpt
    Attpath = CurrentProject.Path & "\Report WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Report WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    Set db = CurrentDb
    Set rst = db.OpenRecordset("EmailContent")
    If rst.RecordCount = 0 Then
        MsgBox "Favor registrar padrão de corpo do email!", vbCritical, "WIP MFPA"
        rst.Close
        Exit Sub
    End If
    rst.MoveFirst
    Do Until rst!ID = 1
        rst.MoveNext
    Loop
    EmailSubj = rst!Assunto
    EmailBody = rst!Corpo
    rst.Close
    Set rst = db.OpenRecordset("EmailList")
    If rst.RecordCount = 0 Then
        MsgBox "Favor cadastrar endereços de email de destinatário!", vbCritical, "WIP MFPA"
        rst.Close
        Exit Sub
    End If
    rst.MoveFirst
    Destination = rst!Email
    Set olApp = CreateObject("Outlook.Application")
    Set olMsg = olApp.CreateItem(olMailItem)
    Do
        rst.MoveNext
        If rst.EOF Then Exit Do
        Destination = Destination & " ; " & rst!Email
    Loop
    olMsg.To = "" & Destination
    olMsg.Subject = EmailSubj
    olMsg.Body = EmailBody
    Set MyAtt = olMsg.Attachments
    MyAtt.Add "" & Attpath
    olMsg.Send
    Set olMsg = Nothing
    Set rst = Nothing
    Set db = Nothing
    MsgBox "Report enviado via Email com sucesso!", vbOKOnly, "WIP MFPA"
    Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & " !", vbCritical, "WIP MFPA"
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
End Sub

Private Sub Report_WIP_Print_Botao_Click()
'macro para chamar a subrotina de criação do report wip e imprimí-lo

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

    Me.Visible = False
    Call WIPRpt
    DoCmd.OpenReport "Report WIP", acViewPreview
    DoCmd.RunCommand acCmdPrint
    DoCmd.Close acReport, "Report WIP"
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Is = 2501
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

Private Sub SAIR_Click()
'macro para sair do app

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
    If MsgBox("Você realmente deseja sair do aplicativo?", vbYesNo, "WIP MFPA") = vbYes Then
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

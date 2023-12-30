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
    Me.Atualizar_Painel_�cone.SetFocus
    EventExec2 = True
    Me.Requery
    Me.Refresh
    EventExec2 = False
End Sub

Private Sub Atualizar_Painel_�cone_Click()
'macro para atualizar os dados do painel
Atualizar_Painel_Click
End Sub

Private Sub Exportar_Excel_Btn_Click()
'macro para exportar o arquivo de relat�rio de produ��o em formato excel

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
    
    Me.Exportar_Excel_�cone.SetFocus
    Do
        InitialDate = InputBox("Insira a data de in�cio: ", "WIP MFPA")
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
    StrSQL = "INSERT INTO Apontamento ( Data, Hora, Opera��o, M�quina, Produzido, Refugado, Refugo ) " & _
            "SELECT Hist�rico.Data, Hist�rico.Hora, Hist�rico.Opera��o, Hist�rico.M�quina, Hist�rico.Produzido, Hist�rico.Refugado, FormatPercent([Refugado]/[Produzido],2) AS Refugo " & _
            "FROM Hist�rico WHERE (((Hist�rico.Data) Between #" & InitialDate & "# And #" & EndDate & "#)) " & _
            "GROUP BY Hist�rico.Data, Hist�rico.Hora, Hist�rico.Opera��o, Hist�rico.M�quina, Hist�rico.Produzido, Hist�rico.Refugado, FormatPercent([Refugado]/[Produzido],2);"
    DoCmd.RunSQL StrSQL
    DoCmd.Hourglass False
    MsgBox "Relat�rio de Produ��o exportado com sucesso!", vbOKOnly, "WIP MFPA"
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Sub

Private Sub Exportar_Excel_�cone_Click()
'macro para exportar o arquivo de relat�rio de produ��o em formato excel
Exportar_Excel_Btn_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'macro para carregar a vers�o do programa

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

Private Sub �cone_Sair_Click()
'macro para sair do app
SAIR_Click
End Sub

Private Sub Quadro_BD_Click()
'macro para sistematizar a altern�nica do modo de visualiza��o do banco de dados entre formul�rio e tabelas
    If Me.Quadro_BD.Value = 1 Then
        Me.Guia_Tabelas.Visible = False
        Me.Guia_Form.Visible = True
    Else
        Me.Guia_Tabelas.Visible = True
        Me.Guia_Form.Visible = False
    End If
End Sub

Private Sub Report_Detalhado_Botao_Click()
'macro para chamar a subrotina de cria��o do report wip detalhado e exib�-lo em modo de visualiza��o de impress�o

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
    Attpath = CurrentProject.Path & "\Relat�rio Detalhado WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Relat�rio Detalhado WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    DoCmd.OpenReport "Relat�rio Detalhado WIP", acViewPreview
    DoCmd.RunCommand acCmdZoom75
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
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
'macro para exportar em pdf o relat�rio de erros e exib�-lo em modo de visualiza��o de impress�o

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
    
    Me.Report_Erro_�cone.SetFocus
    If MsgBox("Deseja atualizar e exportar novo relat�rio de erros do sistema?", vbYesNo, "WIP MFPA") = vbYes Then
        If Backlog.RecordCount = 0 Then
            MsgBox "N�o h� erros para serem exibidos!", vbOKOnly, "WIP MFPA"
        Else
            Me.Visible = False
            Attpath = CurrentProject.Path & "\Relat�rio de Erros.pdf"
            DoCmd.OutputTo acOutputReport, "Relat�rio de Erros", acFormatPDF, Attpath, False, , , acExportQualityPrint
            DoCmd.OpenReport "Relat�rio de Erros", acViewPreview
            DoCmd.RunCommand acCmdZoom75
            StrSQL = "DELETE ErrorBacklog.* FROM ErrorBacklog;"
            DoCmd.RunSQL StrSQL
        End If
    End If
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Sub

Private Sub Report_Erro_�cone_Click()
'macro para exportar em pdf o relat�rio de erros e exib�-lo em modo de visualiza��o de impress�o
Report_Erro_Btn_Click
End Sub

Private Sub Report_WIP_Botao_Click()
'macro para chamar a subrotina de cria��o do report wip e exib�-lo em modo de visualiza��o de impress�o

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
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
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
'macro para salvar o relat�rio em pdf e enviar por email de acordo com a lista de email disponibilizada

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
    Attpath = CurrentProject.Path & "\Relat�rio Detalhado WIP.pdf"
    DoCmd.OutputTo acOutputReport, "Relat�rio Detalhado WIP", acFormatPDF, Attpath, False, , , acExportQualityPrint
    Set db = CurrentDb
    Set rst = db.OpenRecordset("EmailContent")
    If rst.RecordCount = 0 Then
        MsgBox "Favor registrar padr�o de corpo do email!", vbCritical, "WIP MFPA"
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
        MsgBox "Favor cadastrar endere�os de email de destinat�rio!", vbCritical, "WIP MFPA"
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
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
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
'macro para chamar a subrotina de cria��o do report wip detalhado e imprim�-lo

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
    DoCmd.OpenReport "Relat�rio Detalhado WIP", acViewPreview
    DoCmd.RunCommand acCmdPrint
    DoCmd.Close acReport, "Relat�rio Detalhado WIP"
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Is = 2501
            Resume Next
        Case Else
            MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
            Backlog.AddNew
            Backlog!C�digo = Err.Number
            Backlog!Descri��o = Err.Description
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
'macro para salvar o relat�rio em pdf e enviar por email de acordo com a lista de email disponibilizada

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
        MsgBox "Favor registrar padr�o de corpo do email!", vbCritical, "WIP MFPA"
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
        MsgBox "Favor cadastrar endere�os de email de destinat�rio!", vbCritical, "WIP MFPA"
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
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
End Sub

Private Sub Report_WIP_Print_Botao_Click()
'macro para chamar a subrotina de cria��o do report wip e imprim�-lo

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
            Backlog!C�digo = Err.Number
            Backlog!Descri��o = Err.Description
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

    Me.�cone_Sair.SetFocus
    If MsgBox("Voc� realmente deseja sair do aplicativo?", vbYesNo, "WIP MFPA") = vbYes Then
        DoCmd.Quit acQuitSaveAll
    End If
Exit Sub
ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & "! Chamar suporte para consertar o erro no sistema", vbCritical, "WIP MFPA"
    Backlog.AddNew
    Backlog!C�digo = Err.Number
    Backlog!Descri��o = Err.Description
    Backlog!Local = FileName & " / " & LocalObject & " - " & LocalControl
    Backlog!Data = Date
    Backlog!Hora = Time
    Backlog.Update
    Backlog.Close
    Set db = Nothing
    Set Backlog = Nothing
    Resume Next
End Sub

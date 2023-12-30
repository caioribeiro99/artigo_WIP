Attribute VB_Name = "ReportMod"
Option Compare Database
Option Explicit

Sub DelTbl(TblName As String)
'macro para apagar tabela auxiliar (utilizada para geração do relatório)
    On Error Resume Next
    DoCmd.DeleteObject acTable, TblName
End Sub

Function WIPRpt()
'macro para criar o relatório de WIP

Dim StrSQL As String
Dim db As DAO.Database
Dim rst As Recordset, rst2 As Recordset, rst3 As Recordset, rst4 As Recordset, rst5 As Recordset
Dim Current_Oper As String, Current_AN As String
Dim Next_Oper As String
Dim FLUXO As String
Dim TECNOLOGIA As String
Dim Agrupamento As String
Dim IDFluxo As Integer

    DoCmd.Hourglass True
    StrSQL = "DELETE [Report WIP].* FROM [Report WIP];"
    DoCmd.RunSQL StrSQL
    StrSQL = "INSERT INTO [Report WIP] ( Data, OPA, AN, Operação, Quantidade, Pirâmide, Parcial, WIP ) " & _
             "SELECT Fluxo.Data, Fluxo.OPA, Fluxo.AN, Fluxo.Operação, Fluxo.Total AS Quantidade, Fluxo.Pirâmide, Fluxo.Parcial, Date()-[Data] AS WIP " & _
             "FROM Fluxo;"
    DoCmd.RunSQL StrSQL
    Set db = CurrentDb
    Set rst = db.OpenRecordset("Report WIP")
    Set rst2 = db.OpenRecordset("Fluxo por Tecnologia")
    Set rst3 = db.OpenRecordset("Sequência de Operações por Fluxo e Tecnologia")
    Set rst4 = db.OpenRecordset("BD Fluxo por Tecnologia")
    Set rst5 = db.OpenRecordset("BD Agrupamentos")
    rst.MoveFirst
    Do
        Current_AN = rst!AN
        rst4.MoveFirst
            Do Until rst4!AN = Current_AN
                rst4.MoveNext
                If rst4.EOF Then
                    MsgBox Current_AN & " não está cadastrada no sistema!", vbCritical, "WIP MFPA"
                    Exit Do
                End If
            Loop
            If rst4.EOF Then GoTo Final
        FLUXO = rst4!FLUXO
        TECNOLOGIA = rst4!TECNOLOGIA
        Current_Oper = rst!Operação
        rst2.MoveFirst
            Do Until rst2!TECNOLOGIA = TECNOLOGIA And rst2!FLUXO = FLUXO
                rst2.MoveNext
                If rst2.EOF Then
                    Exit Do
                End If
            Loop
        IDFluxo = rst2!ID_Fluxo
        rst3.MoveFirst
            Do Until rst3!ID_Fluxo = IDFluxo And rst3!TECNOLOGIA = TECNOLOGIA And rst3!FLUXO = FLUXO And rst3!Operação = Current_Oper
            rst3.MoveNext
            If rst3.EOF Then
                Exit Do
            End If
            Loop
        If rst3.EOF Then GoTo Final
        rst3.MoveNext
        Next_Oper = rst3!Operação
        rst5.MoveFirst
            Do Until rst5!AN = Current_AN
            rst5.MoveNext
            If rst5.EOF Then
                Exit Do
            End If
            Loop
        Agrupamento = rst5!Agrupamento
        rst.Edit
        rst!Próxima_Operação = Next_Oper
        rst!Agrupamento = Agrupamento
        rst!TECNOLOGIA = TECNOLOGIA
        rst!FLUXO = FLUXO
        rst.Update
Final:
        rst.MoveNext
        If rst.EOF Then
            Exit Do
        End If
    Loop
DoCmd.Hourglass False
End Function

Sub ScriptModelo()

    On Error Resume Next
    
    Dim i  As Long
    Dim iUltimaLinha  As Long
    Dim iPercentualConcluido As Double
    Dim Z  As Long
    Application.ScreenUpdating = False

    Z = 8

    iUltimaLinha = ActiveSheet.Range("A8").End(xlDown).row
    iUltimaLinha = iUltimaLinha - 7
    
    frmBarraDeProgresso.Show False

    TempoInicial = ""
    TempoFianl = ""
    TempoEstimado = ""

    For i = 1 To iUltimaLinha
        If Range("A" & Z) = "" Then
            Exit For
        Else
    
        TempoInicial = Format(Now(), "ss")
        
        'Conexão com o SAP atraves de bibliotecas
        
        Application.DisplayAlerts = False
        Dim sap As New SAP_Manager
        Dim sf As New SAP_Facilites
        Dim xf As New ExcelFacilites
        Set session = sap.getSession(False)

        iPercentualConcluido = i / iUltimaLinha
           
        With frmBarraDeProgresso
            .tempo.Caption = Format(TimeSerial(0, 0, TempoEstimado), "hh:mm:ss")
            .atual.Caption = "Processando  " & Range("A" & Z).Value
            .progresso.Caption = "Parte  " & i & "/" & iUltimaLinha
            .framePb.Caption = Format(iPercentualConcluido, "0%") & " Concluído"
            .progressBar.Width = iPercentualConcluido * (.framePb.Width - 10)
        End With
        DoEvents    'Permite que sejam visualizadas as mudanças nos controles do formulário
        

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

        ' O código da sua macro vai aqui
        
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------- 
 
        TempoFinal = Format(Now(), "ss")
        TempoEstimado = (TempoFinal - TempoInicial) * (iUltimaLinha - i)

         Z = Z + 1

    End If

Next

ThisWorkbook.Save
Unload frmBarraDeProgresso
MsgBox "Processo concluído.", vbInformation, "Informação"

End Sub







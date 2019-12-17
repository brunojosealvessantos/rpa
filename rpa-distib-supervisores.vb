Sub AlterarGerente()

    On Error Resume Next

    Dim Z  As Long
    Dim i  As Long
    Dim iUltimaLinha  As Long
    Dim iPercentualConcluido As Double
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
        
    '------------------------------------------------------------------------------------------------------------------
           
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME22N"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 26
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = Range("A" & Z).Value
        session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").caretPosition = 10
        session.findById("wnd[1]").sendVKey 0
        'session.findById("wnd[0]/tbar[1]/btn[7]").press  
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        For x = 10 To 20
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & x & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11").Select
        Next x
        For x = 10 To 20
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & x & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/btn%#AUTOTEXT001").press
        Next x
        session.findById("wnd[0]/usr/ctxtCI_EKKODB-YYGERENTE").Text = Range("B" & Z).Value
        session.findById("wnd[0]/usr/ctxtCI_EKKODB-YYGERENTE").SetFocus
        session.findById("wnd[0]/usr/ctxtCI_EKKODB-YYGERENTE").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtCI_EKKODB-YYEST_ORIGEM_ACC").Text = "53529100"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        M = session.findById("wnd[0]/sbar").Text
        Range("C" & Z) = M
        TempoFinal = Format(Now(), "ss")
        TempoEstimado = (TempoFinal - TempoInicial) * (iUltimaLinha - i) 
        Z = Z + 1
        
        End If

    Next i

        ThisWorkbook.Save
        Unload frmBarraDeProgresso
        MsgBox "Processo de Transferência de Chave Finalizado."

End Sub



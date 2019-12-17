Sub Arquivador()

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
        
        Elemento = Range("A" & Z).Value  
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME23N"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 26
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = Range("A" & Z).Value
        session.findById("wnd[1]").sendVKey 0
        'session.findById("wnd[0]/tbar[1]/btn[7]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").SelectAll
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton "%ATTA_EXPORT"
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & UsuarioRede & "\Desktop\Base_de_Anexos\" & Elemento
        session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 21 
        For X = 1 To 500
            session.findById("wnd[1]/tbar[0]/btn[11]").press
        Next
        session.findById("wnd[1]").Close
        Range("B" & Z).Value = "Arquivamento Concluido"

        TempoFinal = Format(Now(), "ss")
        TempoEstimado = (TempoFinal - TempoInicial) * (iUltimaLinha - i)
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

        Z = Z + 1

        End If

    Next

ThisWorkbook.Save
Unload frmBarraDeProgresso
MsgBox "Processo concluído.", vbInformation, "Informação"

End Sub

Public Function UsuarioRede()
    Dim GetUserN
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    GetUserN = ObjNetwork.userName
    UsuarioRede = GetUserN
End Function






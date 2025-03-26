Sub Export()

    Windows("MacroFederais.XLSX").Activate
    Workbooks("MacroFederais").Close SaveChanges:=False
    
'Abre sap e Gera o razao
    Dim cel As Range
    Dim cell As Range
    Dim Tempo As Date
    Dim contas As Long
    Dim conta As Double
    Dim Empresa As Range
    Dim Empresa2 As Range
    Dim Empresa3 As Range
    Dim Empresa4 As Range
    Dim Empresa5 As Range
    
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
    
   With Planilha2
    
    Set Empresa3 = Range("B23")
    Set Empresa4 = Range("B24")
    Set Empresa5 = Range("B25")
    Set Empresa = Range("E4")
    Set Empresa2 = Range("E5")
    Set cel = Range("C4")
    Set cell = Range("C5")
    
   End With
   
   conta = 10
   
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "FAGLL03"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = Empresa
    Session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").setFocus
    Session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").caretPosition = 4
    Session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press
    
    With Planilha3
        
        Do
    
        conta = conta + 1
        
        If .Cells(conta, 13).Value <> "" Then
        
        contas = .Cells(conta, 13).Value
    
        End If

Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = contas
Session.findById("wnd[1]/tbar[0]/btn[13]").press

        Loop Until .Cells(conta, 13).Value = ""
            
    End With

Session.findById("wnd[1]/tbar[0]/btn[8]").press
Session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = Empresa3
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = Empresa4
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").Text = Empresa5
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").setFocus
Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").caretPosition = 4
Session.findById("wnd[1]/tbar[0]/btn[8]").press
Session.findById("wnd[0]/usr/radX_AISEL").Select
Session.findById("wnd[0]/usr/ctxtSD_BUKRS-HIGH").Text = Empresa2
Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = cel
Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = cell
Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus
Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 8
Session.findById("wnd[0]/tbar[1]/btn[8]").press
Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "T:\MacroEquipeIndiretos\MacroAutonomos"
Session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
Session.findById("wnd[1]/tbar[0]/btn[12]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press

Tempo = Now + TimeValue("00.00.7")

Call Application.OnTime(Tempo, "GerarINSSAut", , True)

End Sub

Sub GerarINSSAut()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
 
 
Windows("EXPORT.XLSX").Activate
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy Before:=Workbooks("Macro_Automatizada_LEO.xlsm").Sheets(3)
    
Windows("EXPORT.XLSX").Activate
    Workbooks("EXPORT").Close SaveChanges:=False

Windows("Macro_Automatizada_LEO").Activate
    Sheets("Sheet1").Select
    Cells.Select
    Selection.RemoveSubtotal
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$R$1952").AutoFilter Field:=2, Criteria1:=RGB(255, _
        255, 153), Operator:=xlFilterCellColor
    Rows("2:1048576").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    Range("B1949:B1950").Select
    Range("B1950").Activate
    ActiveSheet.Range("$A$1:$R$1939").AutoFilter Field:=2, Criteria1:=RGB(255, _
        255, 0), Operator:=xlFilterCellColor
    Rows("2:1048576").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    Range("B1947:B1948").Select
    Range("B1948").Activate
    Selection.AutoFilter
    
End Sub
Sub Converter()

On Error GoTo Erro
Dim Valor As Double
Dim Linha As Long

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
Linha = 1


With Planilha1
    
    Do
    
        Linha = Linha + 1
        
        If .Cells(Linha, 1).Value <> "" Then
        
        Valor = .Cells(Linha, 1).Value
        
        .Cells(Linha, 1).Value = Valor
    
    End If
    
Loop Until .Cells(Linha, 1).Value = ""

End With
Exit Sub
Erro:
MsgBox "Macro", vbInformation, "Macro Gerada com Sucesso!"

End Sub
Sub Converter2()
    
On Error GoTo Erro
Dim Valor As Double
Dim Linha As Double

Linha = 1


With Planilha1
    
    Do
    
        Linha = Linha + 1
        
        If .Cells(Linha, 2).Value <> "" Then
        
        Valor = .Cells(Linha, 2).Value
        
        .Cells(Linha, 2).Value = Valor
    
    End If
    
Loop Until .Cells(Linha, 2).Value = ""

End With

Exit Sub
Erro:
MsgBox "Macro", vbInformation, "Macro Gerada com Sucesso!"


End Sub


Sub Preencher()
On Error GoTo Erro
Dim Valor2 As Long
Dim Linha As Double
Dim Linha2 As Double
Linha2 = 2
Linha = 1
Valor2 = 1

With Planilha1
    
        Do
        
        Linha = Linha + 1
        Linha2 = Linha2 + 1
        
            If .Cells(Linha, 2).Value = "" Then
            
            .Cells(Linha, 2).Value = Valor2
            
            
            End If
            
        Loop Until .Cells(Linha2, 2).Value <> ""
            
        
End With


Exit Sub
Erro:
MsgBox "Macro", vbInformation, "Macro Gerada com Sucesso!"

End Sub

Sub Tempoparaexecutar()
Tempo = Now + TimeValue("00.00.7")

Call Application.OnTime(Tempo, "Export", , True)


End Sub
 Sub PagamentosAutonomos()

    Dim Ano As Long
    
    With Planilha2
    
    
    Range("G40").Copy
    Range("G39").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    .Cells(39, 7).Value = Format(Date, "YYYY")
    Ano = .Cells(39, 7)
    
    End With
    
    Call Conexao_SAP("ZF0077")
    
    Session.findById("wnd[0]/usr/ctxtSP$00001-LOW").Text = ""
    Session.findById("wnd[0]/usr/txtSP$00002-LOW").Text = ""
    Session.findById("wnd[0]/usr/txtSP$00003-LOW").Text = Ano
    Session.findById("wnd[0]/usr/ctxtSP$00005-LOW").Text = "IC"
    Session.findById("wnd[0]/usr/ctxtSP$00006-LOW").Text = "F0"
    Session.findById("wnd[0]/usr/ctxtSP$00006-LOW").setFocus
    Session.findById("wnd[0]/usr/ctxtSP$00006-LOW").caretPosition = 2
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "T:\MacroEquipeIndiretos\MacroAutonomos"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "MacroFederais.XLSX"
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]").sendVKey 12
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    Session.findById("wnd[0]/usr/btnSTARTBUTTON").press
    
 End Sub
 Sub PreparaRazao()
 
    Range("G40").Copy
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F19").Select
    Selection.NumberFormat = "yyyy"
    
 End Sub
 
    

Sub GerarINSS_AUT()
    PagamentosAutonomos
    Tempoparaexecutar
    
End Sub


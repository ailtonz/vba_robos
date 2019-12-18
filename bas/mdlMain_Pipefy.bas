Attribute VB_Name = "mdlMain_Pipefy"
Private Const ColumnIndex As Integer = 2
Private Const InicioDaPesquisa As Long = 2
Private selectedRange As Range
Private cell As Range
Private x As Long

Sub pipefy_importDate(ByVal control As IRibbonControl)
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets("Pipefy")

ws.Activate
ws.Visible = IIf(ws.Visible = xlSheetVisible, xlSheetHidden, xlSheetVisible)

'
''' Principal
'Dim obj As New clsRdp
''Dim strDominio As String: strDominio = Etiqueta("ServerDominio")
''Dim strCaminho As String: strCaminho = Application.ActiveWorkbook.Path & "\Vms\"
'
''' linhas e colunas
'Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row 'lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
'
''' Confirmar de execução
'Dim sTitle As String:       sTitle = "Abrir Vm's (Selecionadas)"
'Dim sMessage As String:     sMessage = "Deseja abrir a(s) VMs selecionado(s) ?"
'Dim resposta As Variant
'
''' Base de VM's
''CreateDir strCaminho
'
'If (ActiveSheet.Name <> ws.Name) Then
'    ws.Visible = xlSheetVisible
'    ws.Activate
'Else
''    If Len(Dir(strCaminho, vbDirectory)) > 0 Then
'        resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'        If (resposta = vbYes) Then
''            Set selectedRange = Application.Selection
'
'            For Each cell In selectedRange.Cells
'                For x = InicioDaPesquisa To lRow - 1
''                    With obj
''                        If (ws.Range("B" & x).Value = cell.Value And ws.Range("E" & x).Value = "RDP") Then
''                            '' Dados para arqivo
''                            .strAddress = Trim(CStr(ws.Range("B" & x).Value))
''                            .strUserName = strDominio & "\" & CStr(ws.Range("C" & x).Value)
''                            .strUserPass = Trim(CStr(ws.Range("D" & x).Value))
''                            .strPath = IIf(CStr(ws.Range("G" & x).Value) = "", strCaminho, CStr(ws.Range("G" & x).Value))
''                            .strRun = CStr(ws.Range("H" & x).Value)
''                            .gerarRdp
''                            .gerarCredencial
''                            '' Copia de senha
''                            ClipBoardThis Trim(CStr(ws.Range("D" & x).Value))
''                        End If
''                    End With
'                Next x
'            Next cell
'            Set obj = Nothing
'        End If
''    Else
''        MsgBox "Por favor indique o caminho padrão onde será gravado os arquivos RDP.", vbCritical + vbOKOnly, "Atenção: Caminho padrão não foi criado! "
''    End If
'End If

End Sub

Sub epf_exportDate(ByVal control As IRibbonControl)
    MsgBox "EM TESTES!", vbInformation + vbOKOnly, "epf_exportDate"
End Sub


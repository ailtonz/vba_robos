Attribute VB_Name = "mdlMain_RDP"
Private Const ColumnIndex As Integer = 2
Private Const InicioDaPesquisa As Long = 2
Private selectedRange As Range
Private cell As Range
Private x As Long

Sub rdp_importDate(ByVal control As IRibbonControl)
''' Worksheet
'Dim wsDest As Worksheet: Set wsDest = Worksheets("vms")
'Dim wsCopy As Workbook
'Dim Sheet As Worksheet
'
'ws.Activate
'ws.Visible = IIf(ws.Visible = xlSheetVisible, xlSheetHidden, xlSheetVisible)
'
'If (ws.Visible = xlSheetVisible) Then

    MsgBox "EM TESTES!", vbInformation + vbOKOnly, "rdp_importDate"

'End If

'
''' Principal
''Dim strSenha As String: strSenha = Etiqueta("SenhaPadrao")
'
''' linhas e colunas
'Dim lCopyLastRow As Long
'Dim lDestLastRow As Long
'
''' Confirmar de execução
'Dim sTitle As String:       sTitle = "Importar base de VMs"
'Dim sMessage As String:     sMessage = "Deseja importar base de VMs ?"
'Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'
'If (resposta = vbYes) Then
'    Application.ScreenUpdating = False
'
'    Select Case StrPtr(resposta)
'        Case 0
'             MsgBox "Atualização cancelada.", 64, sTitle
'            Exit Sub
'        Case Else
'            FileToOpen = Application.GetOpenFilename _
'                        (Title:="Por favor selecione a planilha para importação de dados", _
'                        FileFilter:="Report Files *.xls* (*.xls*),")
'            If FileToOpen = False Then
'                MsgBox "Nenhum arquivo selecionado.", vbExclamation, "ERROR - Importação de dados"
'                Exit Sub
'            Else
'                Set wsCopy = Workbooks.Open(Filename:=FileToOpen)
'                For Each Sheet In wsCopy.Sheets
'                      lCopyLastRow = Sheet.Cells(Sheet.Rows.Count, "A").End(xlUp).Row
'                      lDestLastRow = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).Offset(1).Row
'                      Sheet.Range("A" & InicioDaPesquisa & ":D" & lCopyLastRow).Copy wsDest.Range("B" & lDestLastRow)
'                      wsDest.Range("A" & lDestLastRow).Value = Sheet.Name
'                Next Sheet
'            End If
'            wsCopy.Close
'    End Select
'
'    Application.ScreenUpdating = True
'End If

End Sub

Sub rdp_open(ByVal control As IRibbonControl)

    MsgBox "EM TESTES!", vbInformation + vbOKOnly, "rdp_open"

''' Worksheet
'Dim ws As Worksheet: Set ws = Worksheets("vms")
'
''' Principal
'Dim obj As New clsRdp
'Dim strDominio As String: strDominio = Etiqueta("ServerDominio")
'Dim strCaminho As String: strCaminho = Application.ActiveWorkbook.Path & "\Vms\"
'
''' linhas e colunas
'Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row 'lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
'
''' Confirmar de execução
'Dim sTitle As String:       sTitle = "Abrir Vm's (Selecionadas)"
'Dim sMessage As String:     sMessage = "Deseja abrir a(s) VMs selecionado(s) ?"
'Dim resposta As Variant
'
'ws.Activate
'ws.Visible = IIf(ws.Visible = xlSheetVisible, xlSheetHidden, xlSheetVisible)
'
'If (ws.Visible = xlSheetVisible) Then
'
'    '' Base de VM's
'    CreateDir strCaminho
'
'    If (ActiveSheet.Name <> ws.Name) Then
'        ws.Visible = xlSheetVisible
'        ws.Activate
'    Else
'        If Len(Dir(strCaminho, vbDirectory)) > 0 Then
'            resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'            If (resposta = vbYes) Then
'                Set selectedRange = Application.Selection
'
'                For Each cell In selectedRange.Cells
'                    For x = InicioDaPesquisa To lRow - 1
'                        With obj
'                            If (ws.Range("B" & x).Value = cell.Value And ws.Range("E" & x).Value = "RDP") Then
'                                '' Dados para arqivo
'                                .strAddress = Trim(CStr(ws.Range("B" & x).Value))
'                                .strUserName = strDominio & "\" & CStr(ws.Range("C" & x).Value)
'                                .strUserPass = Trim(CStr(ws.Range("D" & x).Value))
'                                .strPath = IIf(CStr(ws.Range("G" & x).Value) = "", strCaminho, CStr(ws.Range("G" & x).Value))
'                                .strRun = CStr(ws.Range("H" & x).Value)
'                                .gerarRdp
'                                .gerarCredencial
'                                '' Copia de senha
'                                ClipBoardThis Trim(CStr(ws.Range("D" & x).Value))
'                            End If
'                        End With
'                    Next x
'                Next cell
'                Set obj = Nothing
'            End If
'        Else
'            MsgBox "Por favor indique o caminho padrão onde será gravado os arquivos RDP.", vbCritical + vbOKOnly, "Atenção: Caminho padrão não foi criado! "
'        End If
'    End If
'
'End If

End Sub

Attribute VB_Name = "teste_conceito"

'Public Sub CopySheetToEndAnotherWorkbook()
'    ActiveSheet.Copy After:=Workbooks("Book1.xlsx").Sheets(Workbooks("Book1.xlsx").Worksheets.Count)
'End Sub
'


'Sub teste_CountRows()
'    Dim myCount As Integer
'    myCount = Selection.Rows.Count
'    MsgBox "This selection contains " & myCount & " row(s)", vbInformation, "Count Rows"
'End Sub
'
'Sub RoundToZero3()
' For Each c In ActiveCell.CurrentRegion.Cells
'    Debug.Print c.Value
' Next
'End Sub
'
'Sub teste_split()
'''https://excelmacromastery.com/excel-vba-array/
'
'Dim arr As Variant: arr = Split("James:Earl:Jones", ":")
'
' For Each c In arr
'    Debug.Print c
' Next
'
'End Sub
'
'Sub teste_listarEpf()
'
'Dim arr As Variant
'Dim str01 As String
'Dim str02 As String
'
'
' For Each c In ActiveCell.CurrentRegion.Cells
'    str01 = Split(c.Value, "|")(1)
'    str02 = Split(c.Value, "|")(0)
'    criarEpf str01, str02
' Next
'End Sub
'
'
'Sub teste_loop()
'
'For Each o In Range("b1:b10")
'
'    Debug.Print o.Value
'
'Next o
'
'End Sub
'
'Sub Example()
'    DownloadFile$ = "someFile.ext" 'here the name with extension
'    URL$ = "http://some.web.address/" & DownloadFile 'Here is the web address
'    LocalFilename$ = "C:\Some\Path" & DownloadFile Or CurrentProject.Path & "\" & DownloadFile 'here the drive and download directory
'    MsgBox "Download Status : " & URLDownloadToFile(0, URL, LocalFilename, 0, 0) = 0
'End Sub

'Sub abrirPipefy()
'
'Dim strURL As String: strURL = "https://pipefy.paas.santanderbr.corp/login" 'Etiqueta("Url_Console")
'Dim strUserName As String: strUserName = "ailton.z.da.silva@avanade.com" 'Etiqueta("Url_Console_userName")
'Dim strUserPass As String: strUserPass = "41L70N@@" 'Etiqueta("Url_Console_userPass")
'
'Dim URL As String: URL = strURL
'Dim ObjIE As Object: Set ObjIE = CreateObject("InternetExplorer.Application")
'
'
'With ObjIE
'  .Visible = True
'  .Navigate (URL)
'
''  Do
''    Sleep (50)
''  Loop Until .ReadyState = 4
'
'End With
'
'With ObjIE.document.forms("new_user")
'  .user_login.value = strUserName
'  .user_password.value = strUserPass
'  .submit
'End With
'
'End Sub



'Sub WebPage()
'
'Application.ScreenUpdating = False
'
'    Dim ws As Worksheet: Set ws = Worksheets("Pipefy")
'    Dim WebUrl As String
'    Dim lRow As Long
'
'
'    ws.Activate
''    If (ws.Name = ActiveSheet.Name) Then
''        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
''        For Each cell In ws.Range("$R$2:$R$" & lRow)
''            If (Len(cell.Value) > 0) Then WebUrl = cell.Value
''            Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url " & WebUrl)
''        Next cell
''    End If
'
'Application.ScreenUpdating = True
'
'End Sub


'Sub WebPage()
'
'Application.ScreenUpdating = False
'
'    Dim ws As Worksheet: Set ws = Worksheets("Pipefy")
'    Dim LocalFilename As String: LocalFilename = "c:\temp\"
'    Dim WebUrl As String
'    Dim lRow As Long
'
'
'    ws.Activate
'    If (ws.Name = ActiveSheet.Name) Then
'        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'        For Each cell In ws.Range("$R$2:$R$" & lRow)
'            If (Len(cell.Value) > 0) Then WebUrl = cell.Value
'            'Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url " & WebUrl)
'
'            URLDownloadToFile 0, WebUrl, LocalFilename, 0, 0
'
'        Next cell
'    End If
'
'Application.ScreenUpdating = True
'
'End Sub


'Sub rdp_PathDefalt(ByVal control As IRibbonControl)
'Dim sCaminho As String
'Dim strCaminhoPadrao As String: strCaminhoPadrao = Etiqueta("caminhoRdp")
'Dim sTitle As String:       sTitle = "Caminho padrão"
'Dim sMessage As String:     sMessage = "Deseja alterar o camimho padrão ( " & strCaminhoPadrao & " ) onde ficará salvas as VMs ?"
'Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'
'If (resposta = vbYes) Then
'    Select Case StrPtr(resposta)
'        Case 0
'             MsgBox "Atualização cancelada.", 64, sTitle
'            Exit Sub
'        Case Else
'             sCaminho = GetFolder()
'             If (sCaminho <> "") Then
'                ThisWorkbook.Names("caminhoRdp").Value = sCaminho
'                MsgBox "Caminho atualizado para : " & (sCaminho) & ".", 64, sTitle
'             Else
'                MsgBox "Operação cancelada", vbInformation, sTitle
'             End If
'    End Select
'End If
'End Sub


'Sub criarEpf(strPassword As String, strDestinationFile As String)
'Dim obj As New clsEPF
'Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
'Dim strCaminho As String: strCaminho = Application.ActiveWorkbook.Path & "\"
'Dim strAppName As String: strAppName = "GenEpf.exe"
'
'Dim sTitle As String:       sTitle = "Criar arquivo EPF"
'Dim sMessage As String:     sMessage = "Deseja criar um arquivo EPF ?"
'Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'
'If (resposta = vbYes) Then
'    With obj
'        .strAppPath = strCaminho
'        .strAppName = strAppName
'        .strPassword = strPassword
'        .strDestinationPath = strCaminho
'        .strDestinationFile = strDestinationFile
'        .gerarEpf
'    End With
'Else
'    MsgBox "Operação cancelada", vbInformation, sTitle
'End If
'
'End Sub

''' Sub criarRdpPorSelecao(ByVal control As IRibbonControl)
'Sub criarEpf()
'
'Dim obj As New clsBringTo
'
'Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
'Dim strCaminho As String: strCaminho = Application.ActiveWorkbook.Path & "\"
'Dim strAppName As String: strAppName = "GenEpf.exe"
'
'Dim sTitle As String:       sTitle = "Criar arquivo EPF"
'Dim sMessage As String:     sMessage = "Deseja criar um arquivo EPF ?"
'Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
'
'If (resposta = vbYes) Then
'    With obj
'        .strAppPath = strCaminho
'        .strAppName = strAppName
'        .strPassword = ws.Range("$C$9")
'        .strDestinationPath = strCaminho
'        .strDestinationFile = ws.Range("$C$8")
'        .gerarEpf
'    End With
'Else
'    MsgBox "Operação cancelada", vbInformation, sTitle
'End If
'
'End Sub


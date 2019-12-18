Attribute VB_Name = "mdlMain"
'' [ imageMSO ]
'' https://bert-toolkit.com/imagemso-list.html

'' [ git ]
'' https://githowto.com/pt-BR/create_a_project

Private Const ColumnIndex As Integer = 3
Private Const InicioDaPesquisa As Long = 12
Private Const ColunaTipoDeArquivo As String = "C"
Private Const ColunaStatus As String = "E"
Private Const ColunaTarefa As String = "D"
Private cell As Range
Private strBody As String
Private tmp As String

Option Explicit

Sub create_FileIni(ByVal control As IRibbonControl) '' Criar arquivo INI
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileIni")
Dim strFileName As String: strFileName = ""

'' Principal
Dim o As New cls_file_Ini
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

    '' Carregará apenas itens onde a coluna "Status" estiver diferente de "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value <> "OK" Then
        
            strFileName = cell.Value ''Trim(Split(cell.Value, "|")(0))
        
            With o
                .strFileName = strFileName
                .strFilePath = t
                
                '' Ambiente do robo - "Caminho do Robo"
                .strPathEnvironment = ws.Range(Etiqueta("robot_07_Environment")).Value
                
                .gerarFileIni
            End With
            
            ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK"
        
        End If
        linha = linha + 1
    Next cell

End Sub

Sub create_FileJs(ByVal control As IRibbonControl) '' Criar arquivo Js
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileJs")

'' Principal
Dim o As New cls_file_js
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

'' Consulta a existencia de uma base de dados para carregar funções de controle de fila
Dim dbDados As New Collection: Set dbDados = getData()

    '' Carregará apenas itens onde a coluna "Status" estiver diferente de "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value <> "OK" Then
            
            With o
                .strFileName = cell.Value
                .strFilePath = t
                
                '' Ativar / Desativar a existencia de base(s)
                .CarregarBase = dbDados.Count
                .gerarJs
            End With
            
            ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK"
        End If
        linha = linha + 1
    Next cell

End Sub

Sub create_FileWsf(ByVal control As IRibbonControl) '' Criar arquivo wsf
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileWsf")

Dim strFileName As String: strFileName = ""
Dim strPathFileSource As String: strPathFileSource = ""

'' Principal
Dim o As New cls_file_Wsf
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

'' Consulta a existencia de uma base de dados para carregar funções de controle de fila
Dim dbDados As New Collection: Set dbDados = getData()

    '' Carregará apenas itens onde a coluna "Status" estiver diferente de "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value <> "OK" Then
            
            strFileName = cell.Value ''Trim(Split(cell.Value, "|")(0))
            With o
                .strFileName = strFileName
                .strFilePath = t
                
                '' Carrega todos os procedimentos que serão usados no robo
                .colProcedures = getProcedures
                
                '' Carrega as bibliotecas marcadas na coluna "Status" com "OK"
                .colLibrary = getLibrary
                
                '' Ativar / Desativar a existencia de base(s)
                .CarregarBase = dbDados.Count
                
                '' Ambiente do robo - "Caminho do Robo"
                .strPathEnvironment = ws.Range(Etiqueta("robot_07_Environment")).Value
                
                '' Repositório onde o robo vai coletar a(s) Base(s)
                .strPathRepository = ws.Range(Etiqueta("robot_08_Repository")).Value '"%HOMEDRIVE%%HOMEPATH%\\Downloads\\"
                .gerarWsf
            End With
            
            ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK"
        End If
        linha = linha + 1
    Next cell

End Sub

Sub create_FileEpf(ByVal control As IRibbonControl) '' Criar arquivo epf
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileEpf")
Dim strPassword As String: strPassword = Etiqueta("robot_FileEpf_Password")

'' Principal
Dim o As New clsEPF
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value
Dim pathApp As String: pathApp = Etiqueta("pathApp")
Dim pathName As String: pathName = "GenEpf.exe"

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

'' Confirmar de execução
Dim sTitle As String:       sTitle = "Criar arquivo EPF"
Dim sMessage As String:     sMessage = "Deseja criar um arquivo EPF ?"
Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)

    If (resposta = vbYes) Then
    
        '' Carregará apenas itens onde a coluna "Status" estiver diferente de "OK"
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value <> "OK" Then
                
                With o
                    '' Caminho do app
                    .strAppPath = pathApp
                    '' Aplicação de criptografia
                    .strAppName = pathName
                    '' Password
                    .strPassword = strPassword
                    '' Name file
                    .strDestinationFile = cell.Value
                    '' Caminho onde será salvo o arquivo.epf
                    .strDestinationPath = t
                    .gerarEpf
                End With
                
                ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK"
            End If
            linha = linha + 1
        Next cell
    End If

End Sub

Sub create_BatStart(ByVal control As IRibbonControl) '' Criar bat de start do robo
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)

'' Principal
Dim o As New cls_bat_Start
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value
Dim strFilePathRobo As String: strFilePathRobo = ws.Range(Etiqueta("robot_07_Environment")).Value
Dim strPathRepository As String: strPathRepository = Etiqueta("pathRepository")

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

'' Colecao de filtros
Dim Col As New Collection
Col.add Etiqueta("robot_FileWsf")
Col.add Etiqueta("robot_FileIni")
Col.add Etiqueta("robot_FileEpf")

Dim item(2) As String
Dim x As Integer: x = 0
Dim filtro As Variant

'' Consulta a existencia de uma base de dados para carregar funções de controle de fila
Dim dbDados As New Collection: Set dbDados = getData()

    '' Coleta de dados
    For Each filtro In Col
        linha = InicioDaPesquisa
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, filtro) <> 0 Then
                item(x) = cell.Value
                x = x + 1
            End If
            linha = linha + 1
        Next cell
    Next filtro
    
    '' Criação do arquivo
    With o
        .strFileNameBat = "Exec_" & Left(item(0), Len(item(0)) - 4) & ".bat"
        .strFileNameWsf = item(0)
        .strFileNameIni = item(1)
        .strFileNameEPF = item(2) ''Left(Split(item(2), "|")(0), Len(Split(item(2), "|")(0)) - 4)
        .strFilePathBat = t
        .strFilePathRobo = strFilePathRobo
        .strPathRepository = strPathRepository
        
        '' Ativar / Desativar a existencia de base(s)
        .CarregarBase = dbDados.Count
        
        '' Quantidade padrao
        .strQtdeMassa = "10"
        
        .gerarBatStart
    End With

End Sub

Sub create_BatBkp(ByVal control As IRibbonControl) '' Criar bat de bkp
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)

'' Principal
Dim o As New cls_Bat_Bkp
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value
Dim sFileName As String: sFileName = ws.Range(Etiqueta("robot_batBkp")).Value

    With o
        .strFileName = Split(ws.Range(Etiqueta("robot_07_Environment")).Value, "\")(0)
        .strFilePath = t
        .gerarBatBkp
    End With

End Sub

Sub create_FileRdp(ByVal control As IRibbonControl) '' Criar arquivo rdp
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileRdp")
Dim strDominio As String: strDominio = Etiqueta("ServerDominio")

'' Principal
Dim o As New clsRdp
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value
Dim strFilePathRobo As String: strFilePathRobo = ws.Range(Etiqueta("robot_07_Environment")).Value
Dim strPathRepository As String: strPathRepository = Etiqueta("pathRepository")

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

    '' Carregará apenas itens onde a coluna "Status" estiver diferente de "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value <> "OK" Then
            
            With o
                '' Dados para arqivo
                .strAddress = Trim(Split(cell.Value, "|")(0))
                .strUserName = strDominio & "\" & CStr(Split(cell.Value, "|")(1))
                .strUserPass = Trim(Split(cell.Value, "|")(2))
                .strPath = t
                
                .gerarRdp
                .gerarCredencial
                '' Copia de senha
                ClipBoardThis Trim(.strUserPass)
                
            End With
            
            ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK"
    
        End If
        linha = linha + 1
    Next cell

End Sub


Sub import_pipefy(ByVal control As IRibbonControl)

MsgBox "pipefy"

End Sub

Sub import_vms(ByVal control As IRibbonControl)

MsgBox "vms"

End Sub

Sub import_epf(ByVal control As IRibbonControl)

MsgBox "epf"

End Sub

Sub send_communication_current(ByVal control As IRibbonControl) '' Enviar e-mail com posição da tarefa atual
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim eMail As New clsOutlook

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("eMail_Search")
Dim strTo As String: strTo = Etiqueta("eMail_To")
Dim strCC As String: strCC = Etiqueta("eMail_CC")
Dim strSubject As String: strSubject = Etiqueta("eMail_Subject")

'' Confirmação de envio de e-mail
Dim sTitle As String:       sTitle = ws.Range(Etiqueta("robot_02_Name")).Value
Dim sMessage As String:     sMessage = "Deseja enviar e-mail com posição da tarefa atual ?"
Dim resposta As Variant

'' criar tmp_file apenas para apresentacao
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
If (Dir(pathExit) <> "") Then Kill pathExit

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

    strBody = ""
                    
    With eMail
    
        '' To
        .strTo = strTo
        .strCC = strCC
        .strSubject = strSubject
        
        '' Subject
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaStatus & "$" & linha).Value, ws.Range(strFiltro).Value) <> 0 Then
                If (Len(cell.Value) > 0) Then strBody = strBody & cell.Value & vbNewLine
            End If
            linha = linha + 1
        Next cell
        
        '' Send
        .strBody = strBody
        
        '' Apresentação
        TextFile_Append pathExit, strTo & vbNewLine
        TextFile_Append pathExit, strCC & vbNewLine
        TextFile_Append pathExit, ws.Range(strSubject).Value & vbNewLine
        TextFile_Append pathExit, strBody
        Shell "notepad.exe " & pathExit, vbMaximizedFocus
        Kill pathExit
                
        resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
        If (resposta = vbYes) Then
            .EnviarEmail
            MsgBox "Concluido!", vbInformation + vbOKOnly, sTitle
        End If
        
    End With

End Sub

Sub open_List_Tasks(ByVal control As IRibbonControl) '' listar tarefas
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim cSaida As New Collection

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("eMail_Search")

'' criar tmp_file apenas para apresentacao
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
If (Dir(pathExit) <> "") Then Kill pathExit

Dim t As Variant
Dim tmp As String: tmp = ""

    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            If (ws.Name <> Etiqueta("wbk_Modelo") And ws.Name <> Etiqueta("wbk_vms")) Then
                cSaida.add ws.Name & vbTab & ws.Range(strFiltro).Value & vbNewLine
            End If
        End If
    Next
    
    For Each t In cSaida
        tmp = tmp & t
    Next
    
    TextFile_Append pathExit, tmp
    
    ClipBoardThis tmp
    
    MsgBox "O conteudo tambem foi copiado para o ClipBoard ", vbInformation + vbOKOnly, "Concluido!"
    
    Shell "notepad.exe " & pathExit, vbMaximizedFocus

    Kill pathExit

End Sub

Sub open_Repository(ByVal control As IRibbonControl) '' Criar repositorio
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value '& ActiveSheet.Name
Dim strTemp As String: strTemp = Etiqueta("nameFolders")
Dim item As Variant

    '' BASE
    CreateDir t
    
    For Each item In Split(strTemp, "|")
        CreateDir t & "\" & item
    Next
    
    
    Shell "explorer.exe " + t, vbMaximizedFocus

End Sub

Sub show_Help(ByVal control As IRibbonControl) '' Listar funções da aplicação

    ShowVersion

End Sub

Function getData() As Collection '' Validar existencia de base de dados
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim var As New Collection

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileDb")

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa
    
    '' Carregará apenas itens onde a coluna "Tarefa" estiver diferente de "vazio"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And cell.Value <> "" Then
            var.add cell.Value
        End If
        linha = linha + 1
    Next cell
    
    Set getData = var
End Function

Function getProcedures() As Collection '' Coleção de procedimentos
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim var As New Collection

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileJs")

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa
    
    '' Carregará apenas itens onde a coluna "Status" estiver igual a "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK" Then
            var.add cell.Value
        End If
        linha = linha + 1
    Next cell
    
    Set getProcedures = var
End Function

Function getLibrary() As Collection '' Coleção de bibliotecas
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim var As New Collection

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("robot_FileLibrary")

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa
    
    '' Carregará apenas bibliotecas marcadas na coluna "Status" com "OK"
    For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
        If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, strFiltro) <> 0 And ws.Range("$" & ColunaStatus & "$" & linha).Value = "OK" Then
            var.add cell.Value
        End If
        linha = linha + 1
    Next cell
    
    Set getLibrary = var
End Function

Function ADM_List_Itens(ByVal control As IRibbonControl) '' Listar itens do robo atual para ajuda no copy/paste na console
'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim t As String: t = ws.Range(Etiqueta("robot_06_Help_name")).Value
    
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa
Dim strFiltro As String
Dim notas As New clsNotas

'' Colecao de filtros
Dim colNotas As New clsNotas
Dim recNotas As clsNotas
Dim filtro As clsNotas

    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileWsf")
        .strNotes = t
        .strPath = ws.Range(Etiqueta("robot_07_Environment")).Value
        colNotas.add recNotas
    End With

    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileJs")
        .strNotes = t
        .strPath = ws.Range(Etiqueta("robot_07_Environment")).Value
        colNotas.add recNotas
    End With
    
    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileIni")
        .strNotes = t
        .strPath = ws.Range(Etiqueta("robot_07_Environment")).Value
        colNotas.add recNotas
    End With

    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileEpf")
        .strNotes = "Credenciais\TFC\"
        colNotas.add recNotas
    End With


    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileRdp")
        colNotas.add recNotas
    End With

    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_FileLibrary")
        colNotas.add recNotas
    End With

    Set recNotas = New clsNotas
    With recNotas
        .strName = Etiqueta("robot_Users")
        colNotas.add recNotas
    End With
           
        
    '' criar tmp_file apenas para apresentacao
    Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
    If (Dir(pathExit) <> "") Then Kill pathExit
        
    For Each filtro In colNotas.Itens
        linha = InicioDaPesquisa
        '' Coleta e saida de pesquisa
        TextFile_Append pathExit, vbNewLine & "[" & filtro.strName & "]" & vbNewLine
        
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaTipoDeArquivo & "$" & linha).Value, filtro.strName) <> 0 Then
                TextFile_Append pathExit, filtro.strNotes & cell.Value & vbNewLine & filtro.strPath
            End If
            linha = linha + 1
        Next cell
    
    Next filtro
       
    Shell "notepad.exe " & pathExit, vbMaximizedFocus

    Kill pathExit
    
End Function

Sub ADM_create_debugger(ByVal control As IRibbonControl) '' Criar wsf para testes genericos
Dim t As String: t = CreateObject("WScript.Shell").SpecialFolders("Desktop") ' Etiqueta("pathRepository")
Dim o As New cls_file_debugger
    
    '' Wsf
    With o
        .strFileName = "debugger.wsf"
        .strFilePath = t
        .gerarWsf
    End With
    
    Shell "notepad.exe " & t & "\debugger.wsf", vbMaximizedFocus

End Sub

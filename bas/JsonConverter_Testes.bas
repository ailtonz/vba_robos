Attribute VB_Name = "JsonConverter_Testes"
'''' DESABILITADO: MOTIVO:: NÃO ESTOU USANDO AGORA
'''' DESABILITADO: MOTIVO:: NÃO ESTOU USANDO AGORA
'''' DESABILITADO: MOTIVO:: NÃO ESTOU USANDO AGORA
'''' DESABILITADO: MOTIVO:: NÃO ESTOU USANDO AGORA

'Sub excelToJsonExample()
'Dim jsonItems As New Collection
'Dim jsonDictionary As New Dictionary
'Dim jsonFileObject As New FileSystemObject
'Dim jsonFileExport As TextStream
'
'Dim i As Long
'Dim cell As Variant
'
'Sheets("vms").Activate
'
'Set rng = Range("A2:c129")
'
'For i = 2 To rng.Count
'    jsonDictionary("Tipo") = Cells(i, 1).Value
'    jsonDictionary("VM") = Cells(i, 2).Value
'    jsonDictionary("Login") = Cells(i, 3).Value
'    jsonDictionary("Senha") = Cells(i, 4).Value
'
'    jsonItems.Add jsonDictionary
'    Set jsonDictionary = Nothing
'
'Next i
'
'Set jsonFileExport = jsonFileObject.CreateTextFile("c:\Temp\jsonExample.json", True)
'jsonFileExport.WriteLine (JsonConverter.ConvertToJson(jsonItems, Whitespace:=3))
'
'End Sub

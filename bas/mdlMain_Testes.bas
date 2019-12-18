Attribute VB_Name = "mdlMain_Testes"
Private Sheet As Worksheet
Private cell As Range
Private tmp As String

Sub carregar_Itens()
'' Worksheet
Dim wsCopy As Workbook
            
Application.ScreenUpdating = False

FileToOpen = Application.GetOpenFilename _
            (Title:="Por favor selecione a planilha para importação de dados", _
            FileFilter:="Report Files *.xls* (*.xls*),")
If FileToOpen = False Then
    MsgBox "Nenhum arquivo selecionado.", vbExclamation, "ERROR - Importação de dados"
    Exit Sub
Else
    Set wsCopy = Workbooks.Open(Filename:=FileToOpen)
    
    'ActiveWindow.Visible = False
    
    For Each Sheet In wsCopy.Sheets
        tmp = tmp + vbNewLine + Sheet.Name
    Next Sheet
    
    'ActiveWindow.Visible = True
    
    wsCopy.Close
End If

ClipBoardThis tmp

Application.ScreenUpdating = True

End Sub


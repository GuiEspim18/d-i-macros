Public month As String
Public bd As Variant
Public pwd As Variant

Sub Run()
    ' Declarando a variável ws que vai segurar a tabela atual
    Dim ws As Worksheet
    
    ' Atribuindo à varíavel ws o valor da tabela atual
    Set ws = ThisWorkbook.Sheets("AUTOMATION")
    
    ' Pegando o valor do mês selecionado
    month = ws.Range("B1").Value
        
    RunOKRs
    
End Sub

Sub SelectFile(a As Integer)

    ' Declarando a variável selectedFile que vai guardar o valor em string do caminho do arquivo selecionado
    Dim selectedFile As Variant
    
    ' Declarando a variável wb que guardará o valor do arquivo lido no excel
    Dim wb As Workbook
    
    ' Declarando a variável ws que guardará o valor da tabela que queremos ler do workbook do excel
    Dim ws As Worksheet
    
    ' Declarando a variável dataRange que guardará o Range de valores da ws
    Dim dataRange As Range
    
    ' Declarando a variável data que guardará os ranges de dataRange
    Dim data As Variant
    
    ' Declarando a variável maxRow que guardará o valor máximo de linhas da ws
    Dim maxRow As Long
    
    ' Declarando a variável max col que guardará o valor máximo de colunas da ws
    Dim maxCol As Long
    
    ' Atribuindo à variável selectedFile o valor resultado da seleção de arquivos
    selectedFile = Application.GetOpenFilename( _
        FileFilter:="Arquivos Excel (*.xlsx), *.xlsx", _
        Title:="Selecione um arquivo Excel")
    
    ' Verficando se o arquivo foi seleiconado
    If selectedFile <> False Then
    
        ' Atribuindo ao wb o valor resultado que se obteve ao abrir o arquivo selecioado
        Set wb = Workbooks.Open(selectedFile)
        
        ' Atribuindo a ws o valor da primeira sheet desse arquivo selecioado
        Set ws = wb.Sheets(1) ' Alterar para a planilha desejada, se necessário
        
        ' Encontra a última linha e coluna preenchida
        maxRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        maxCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Define o intervalo dos dados
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(maxRow, maxCol))
        
        ' Armazena os dados em uma matriz
        data = dataRange.Value
        
        ' Armazena os dados na variável correta com base no valor de 'a'
        If a = 0 Then
            bd = data
        ElseIf a = 1 Then
            pwd = data
        End If
        
        ' Fecha o arquivo após leitura
        wb.Close SaveChanges:=False
        
        ' Imprimindo um aviso ao usuário de que se selecionou o arquivo corretamente
        MsgBox "Dados armazenados com sucesso."
    Else
        ' Imprimindo um aviso ao usuário de que não foi selecionado nenhum arquivo
        MsgBox "Nenhum arquivo foi selecionado."
    End If
End Sub


Sub SelectBD()
    SelectFile 0
End Sub


Sub SelectPWD()
    SelectFile 1
End Sub


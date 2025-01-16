Public Function DataNest(ws As Worksheet) As Variant
    ' Inicializando a variável dataRange que vai guardar a matriz
    Dim dataRange As Range
    
    ' Inicializando a variável data que vai guardar a linha da matriz
    Dim data As Variant
    
    ' Inicializando a variável firtsRow que vai guardar a primeira linha preenchida
    Dim firstRow As Long
    
    ' Inicializando a variável firstCol que vai guardar a primeira coluna preenchida
    Dim firstCol As Long
    
    ' Inicializando a variável firtRow que vai guardar a ultima linha preenchida
    Dim lastRow As Long
    
    ' Inicializando a variável lastCol que vai guardar a ultima coluna preenchida
    Dim lastCol As Long
    
    ' Encontrando a primeira linha preenchida
    firstRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, _
                             SearchDirection:=xlNext).Row
                             
    ' Encontrando a primeira coluna preenchida
    firstCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByColumns, _
                             SearchDirection:=xlNext).Column
    
    ' Encontrando a última linha preenchida
    lastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious).Row
    
    ' Encontrando a última coluna preenchida
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious).Column

    ' Definindo o intervalo dos dados
    Set dataRange = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
    
    ' Armazenando os dados em uma matriz
    data = dataRange.Value
    
    ' Retornando a matriz de dados
    DataNest = data
End Function

Public Function GetColumn(dn As Variant, columnName As String) As Variant
    ' Inicializando a variável que vai guardar os valores da coluna selecionada
    Dim columnData() As Variant
    
    ' Inicializando a variável que vai guardar o valor do índice da coluna
    Dim columnIndex As Long
    
    ' Inicializando a variável que vai guardar o valor do número de linhas que a tabela possui
    Dim numRows As Long
    
    ' Inicializando a variável de controle do loop
    Dim i As Long
    
    ' Inicializando a vairável booleana para verificar se a coluna foi encontrada
    Dim found As Boolean
    
    ' Atribuindo à variável numRows o número de linhas que a tabela tem
    numRows = UBound(dn, 1)
    
    ' Redimencionando a variável columnData para comportar o número de linhas
    ReDim columnData(1 To numRows)
    
    ' Atribuindo ao valor found como False
    found = False
    
    ' Loop para achar a coluna
    For i = LBound(dn, 2) To UBound(dn, 2)
        
        ' Se o nome da coluna for igual ao nome passado por parametro
        If dn(1, i) = columnName Then
        
            ' O índice da coluna é igual ao índice da coluna da tabela
            columnIndex = i
            
            ' Atribuindo o valor de True pois a coluna foi encontrada
            found = True
            Exit For
        End If
    Next i
    
    ' Se a coluna não for encontrada
    If Not found Then
        MsgBox "Coluna '" & columnName & "' não encontrada."
        Exit Function
    End If
    
    ' Colocando os valores da coluna no array
    For i = 2 To numRows
        columnData(i - 1) = dn(i, columnIndex)
    Next i
    
    ' Retornando os valores da coluna
    GetColumn = columnData
    
End Function


Public Function GetNumberOfRows(a As Variant) As Long
    ' Verificando se a matriz é vazia
    If IsEmpty(a) Then
        ' Se a matriz for vazia retornar o valor 0
        GetNumberOfRows = 0
    Else
        ' Se a matriz não for vazia retornar o valor de linhas da matriz
        GetNumberOfRows = UBound(a, 1) - 1
    End If
End Function

Public Function GetNumberOfColumns(a As Variant) As Long
    ' Verificando se a matriz é vazia
    If IsEmpty(a) Then
        ' Se a matriz for vazia retornar o valor 0
        GetNumberOfColumns = 0
    Else
        ' Se a matriz não for vazia retornar o valor de linhas da matriz
        GetNumberOfColumns = UBound(a, 2)
    End If
End Function

Public Function GetColumns(a As Variant) As Variant
    ' Variável para pegar o número de colunas da tabela
    Dim numCols As Long
    
    ' Variável para guardar os nomes das colunas
    Dim columns() As String
    
    ' Variável de controle do loop
    Dim i As Long
    
    ' Verificando se a matriz é vazia
    If IsEmpty(a) Then
        ' Se a matriz for vazia a função irá retornar um array vazio e fechará a função
        GetColumns = Array()
        Exit Function
    End If
    
    ' Pegando o número de colunas que a matriz possui
    numCols = UBound(a, 2)
    
    ' Redimensionando o array para armazenar os nomes das colunas
    ReDim columns(1 To numCols)
    
    ' Armazenando os nomes das colunas
    For i = 1 To numCols
        columns(i) = a(1, i)
    Next i
    
    ' Retornando os nomes das colunas
    GetColumns = columns
End Function


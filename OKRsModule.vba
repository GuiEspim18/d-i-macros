Private size As Long
Private names As Variant
Private okr As Variant
Private columns As Variant
Private months As Variant

Public Sub OKRs()
    okr = DataNest(ThisWorkbook.Sheets("OKRs"))
    size = GetNumberOfRows(okrs)
    columns = GetColumns(okrs)
    names = GetColumn(okrs, "Country")
    InitializeMonths
    Calculate
End Sub

Private Function Calculate()
    ' Calculate Okrs
End Function

Private Function InitializeMonths()
    Dim month As Object
    Dim i As Integer
    
    ReDim months(1 To 12)
    
    ' Inicializar os meses
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "January"
    month.Add "original_name", "jan"
    Set months(1) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "February"
    month.Add "original_name", "fev"
    Set months(2) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "March"
    month.Add "original_name", "mar"
    Set months(3) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "April"
    month.Add "original_name", "abr"
    Set months(4) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "May"
    month.Add "original_name", "mai"
    Set months(5) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "June"
    month.Add "original_name", "jun"
    Set months(6) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "July"
    month.Add "original_name", "jul"
    Set months(7) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "August"
    month.Add "original_name", "aug"
    Set months(8) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "September"
    month.Add "original_name", "sep"
    Set months(9) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "October"
    month.Add "original_name", "oct"
    Set months(10) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "November"
    month.Add "original_name", "nov"
    Set months(11) = month
    
    Set month = CreateObject("Scripting.Dictionary")
    month.Add "name", "December"
    month.Add "original_name", "dec"
    Set months(12) = month
End Function

Attribute VB_Name = "Module1"
Type kutu
    en As Integer
    boy As Integer
    yükseklik As Integer
    id As Integer
End Type
Sub importDataFromAnotherWorkbook()

    Dim ws As Worksheet
    Dim filter As String
    Dim targetWorkbook As Workbook, wb As Workbook
    Dim Ret As Variant

    Set targetWorkbook = Application.ActiveWorkbook

    filter = "Text files (*.csv),*.csv"
    Caption = "Please Select an input file "
    Ret = Application.GetOpenFilename(filter, , Caption)

    If Ret = False Then Exit Sub

    Set wb = Workbooks.Open(Ret)

    wb.Sheets(1).Move After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)

    ActiveSheet.Name = "ImportData"
    
    Sheets("ImportData").Range("A1:D100").Copy Destination:=Sheets("Sheet1").Range("B4:E100")
    
    Worksheets("Sheet1").Activate
    
End Sub


Sub deneme()

Dim boyutlar As Range
Dim boyutlar2() As Integer
Dim kalan_boy As Integer
Dim kontey_kalan_en As Integer
Dim layer_kalan_en As Integer
Dim kalan_yükseklik As Integer
Dim belirleyici_en As Integer
Dim layer_kalan_yükseklik() As Integer
ReDim layer_kalan_yükseklik(50)
Dim myfile As String
Dim cellValue As Variant

Set boyutlar = Range(Range("b4"), Range("b4").End(xlDown).End(xlToRight))
m = boyutlar.Rows.Count
n = boyutlar.Columns.Count
ReDim boyutlar2(m, n)

Cells(3, 11) = m
Dim kutular() As kutu
ReDim kutular(m)

Dim kalanlar() As Integer
ReDim kalanlar(m)

For i = 1 To m
    For j = 1 To n
        boyutlar2(i, j) = boyutlar(i, j)
    Next
Next

For i = 1 To m
    kutular(i).en = boyutlar2(i, 1)
    kutular(i).boy = boyutlar2(i, 2)
    kutular(i).yükseklik = boyutlar(i, 3)
    kutular(i).id = boyutlar(i, 4)
Next

kontey_kalan_en = Cells(4, 7)
kalan_boy = Cells(4, 8)
kalan_yükseklik = Cells(4, 9)


For i = 1 To m
    If kutular(i).en <> 0 Then
        kalanlar(i) = kutular(i).id
    End If
Next

enb_yükseklik = 0
number_layer = 0

For g = 1 To m
    kontey_kalan_en = kontey_kalan_en - kutular(g).en
    If kontey_kalan_en > 0 Then
        If kalanlar(g) <> 0 Then
            belirleyici_en = kutular(g).en
            
            Cells(g + 2, g + 14) = kutular(g).id
            'Adding the id of box to the cell means loading
            
            cellValue = kutular(g).id
            kalan_boy = kalan_boy - kutular(g).boy
            kutular(g).en = 0
            kutular(g).boy = 0
            layer_kalan_yükseklik(g) = kutular(g).yükseklik
            kutular(g).yükseklik = 0
            kalanlar(g) = 0
            For j = 1 To m
                If kalan_boy >= kutular(j).boy And kutular(j).boy <> 0 Then
                        Cells(j + 2, g + 14) = kutular(j).id
                        cellValue = kutular(j).id
                        kalan_boy = kalan_boy - kutular(j).boy
                        layer_kalan_en = belirleyici_en - kutular(j).en
                        kutular(j).en = 0
                        kutular(j).boy = 0
                        layer_kalan_yükseklik(j) = kutular(j).yükseklik
                        kutular(j).yükseklik = 0
                        kalanlar(j) = 0
                        For p = 1 To m
                            If layer_kalan_en >= kutular(p).en And kutular(p).en <> 0 Then
                                    Cells(p + 2, g + 14) = kutular(p).id
                                    cellValue = kutular(p).id
                                    layer_kalan_en = layer_kalan_en - kutular(p).en
                                    kutular(p).en = 0
                                    kutular(p).boy = 0
                                    layer_kalan_yükseklik(p) = kutular(p).yükseklik
                                    kutular(p).yükseklik = 0
                                    kalanlar(p) = 0
                            End If
                        Next
                End If
            Next
            For i = 1 To m
                If kutular(i).en <> 0 Then
                        kalanlar(i) = kutular(i).id
                End If
            Next
            
        'If it is impossible to add a new one to the current layer,
        'it should be passed by adding new layer
        number_layer = number_layer + 1
        End If
    End If
    For z = 1 To m
        If layer_kalan_yükseklik(z) > enb_yükseklik And layer_kalan_yükseklik(z) <> 0 Then
                enb_yükseklik = layer_kalan_yükseklik(z)
        End If
    Next
    kalan_boy = Cells(4, 8)
    kalan_yükseklik = kalan_yükseklik - enb_yükseklik
    If kalan_yükseklik > 0 Then
        For w = 1 To m
            If kalanlar(w) <> 0 Then
                If kutular(kalanlar(w)).yükseklik <= kalan_yükseklik Then
                    Cells(w, w + 12) = kutular(w).id
                    cellValue = kutular(w).id
            
                    kalan_boy = kalan_boy - kutular(w).boy
                    kutular(w).en = 0
                    kutular(w).boy = 0
                    kutular(w).yükseklik = 0
                    kalanlar(w) = 0
                    For a = 1 To m
                        If kalan_boy >= kutular(a).boy And kutular(a).boy <> 0 And kalan_yükseklik >= kutular(a).yükseklik Then
                            Cells(a, w + 12) = kutular(a).id
                            cellValue = kutular(a).id
                            kalan_boy = kalan_boy - kutular(a).boy
                            layer_kalan_en = belirleyici_en - kutular(a).en
                            kutular(a).en = 0
                            kutular(a).boy = 0
                            layer_kalan_yükseklik(a) = kutular(a).yükseklik
                            kutular(a).yükseklik = 0
                            kalanlar(a) = 0
                            For b = 1 To m
                                If layer_kalan_en >= kutular(b).en And kutular(b).en <> 0 Then
                                    Cells(b, w + 12) = kutular(b).id
                                    cellValue = kutular(b).id
                                    layer_kalan_en = layer_kalan_en - kutular(b).en
                                    kutular(b).en = 0
                                    kutular(b).boy = 0
                                    layer_kalan_yükseklik(b) = kutular(b).yükseklik
                                    kutular(b).yükseklik = 0
                                    kalanlar(b) = 0
                                End If
                            Next
                        End If
                    Next
                    For t = 1 To m
                        If kutular(t).en <> 0 Then
                            kalanlar(t) = kutular(t).id
                        End If
                    Next
                kalan_yükseklik = kalan_yükseklik - kutular(w).yükseklik
                End If
            End If
        Next
    End If
    Cells(3, 13) = number_layer
Next

Cells(3, 13) = number_layer

Close #1

End Sub





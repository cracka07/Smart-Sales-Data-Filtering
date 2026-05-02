VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub txtFrom_Change()

FilterDate
End Sub
Private Sub txtTo_Change()

FilterDate
End Sub

Private Sub txtProduct_Change()
    Dim tbl As ListObject
    Set tbl = Sheets("SMART_SALES").ListObjects("sales")
    
    tbl.Range.AutoFilter field:=4, Criteria1:="*" & Me.txtProduct & "*"

    If Me.txtProduct = "" Then
    tbl.Range.AutoFilter field:=4
    End If

End Sub

Private Sub txtSalesPerson_Change()

        Range("C10").AutoFilter field:=3, Criteria1:="*" & Me.txtSalesPerson & "*"
        
    If Me.txtSalesPerson = "" Then
        Range("C10").AutoFilter field:=3
    End If

End Sub

Private Sub txtStore_Change()
    Dim lastRow As Long
    Dim sum As Double
    Dim maxSale   As Double
    Dim i As Long
    Dim count As Long
    Dim topProduct As String
    Dim ws As Worksheet
    Set ws = Worksheets("SMART_SALES")
        sum = 0
        maxSale = 0
        Range("C10").AutoFilter field:=2, Criteria1:="*" & Me.txtStore & "*"
     
        If Me.txtStore = "" Then
             Range("C10").AutoFilter field:=2
         End If
         

    
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    sum = Application.WorksheetFunction.Subtotal(109, Range("G10:G" & lastRow))
    count = Application.WorksheetFunction.Subtotal(103, Range("G10:G" & lastRow))
   

For i = 10 To lastRow

    If Rows(i).Hidden = False Then
        
        If Cells(i, 7).Value > maxSale Then
            maxSale = Cells(i, 7).Value
            topProduct = Cells(i, 6).Value
        End If
        
    End If

Next i
    Range("T11").Value = topProduct
   Range("T12").Value = sum
   Range("T13").Value = count
        
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
   
End Sub
'========================================================
' Project: Smart Sales Data Filtering System
' Author: Mariano Ferrer
' Role: Excel VBA Developer
' Date: 2026
'
' Copyright (c) 2026 Mariano Ferrer
'
' Description:
' Excel VBA system designed to dynamically filter sales data
' by date, store, salesperson, and product. The solution
' enables fast exploration and analysis of sales datasets
' directly within Excel using automated filtering logic.
'
' Repository:
' https://github.com/cracka07
'========================================================

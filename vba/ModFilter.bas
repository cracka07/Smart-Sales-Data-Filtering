Attribute VB_Name = "ModFilter"
Option Explicit

Sub FilterDate()
  
   Range("C10").CurrentRegion.AutoFilter field:=1, _
    Criteria1:=">=" & Sheet1.txtFrom, Operator:=xlAnd, _
    Criteria2:="<=" & Sheet1.txtTo
    
        
      If Sheet1.txtFrom = "" Then
         Range("C10").CurrentRegion.AutoFilter field:=1
      End If

   If Sheet1.txtTo = "" Then
         Range("C10").CurrentRegion.AutoFilter field:=1
    End If

End Sub

Sub ShowAll()
    
     Dim tbl As ListObject
    Set tbl = Sheets("SMART_SALES").ListObjects("sales")
  
     tbl.Range.AutoFilter
    
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

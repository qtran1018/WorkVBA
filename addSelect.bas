Attribute VB_Name = "addSelect"
Option Private Module
Sub addPageSelectPage() ' For SP

ActiveWorkbook.Sheets.Add After:=ActiveSheet
Worksheets("Qualified_R&V_Leads").Activate


End Sub

Sub addPageSelectPageOR() ' For OReg

Sheets.Add After:=Worksheets("Leads") 'Makes it work only with a "Leads" tab instead of Activesheet
Worksheets("Leads").Activate


End Sub


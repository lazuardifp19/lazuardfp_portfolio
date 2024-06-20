VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SourceData 
   Caption         =   "Source Data Description"
   ClientHeight    =   6945
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13785
   OleObjectBlob   =   "SourceData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SourceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowSourceData()
    
    Me.Show
    
End Sub

Private Sub UserForm_Initialize()
    
    'Making sure the textboxes are not able to be edited
    txtSource.Enabled = False
    txtConduct.Enabled = False
    txtPeriod.Enabled = False
    txtRegion.Enabled = False
    txtPublished.Enabled = False
    txtDate.Enabled = False
    
    'Multi line textbox for the Original Source
    txtOriginal.Enabled = False
    txtOriginal.MultiLine = True
    txtOriginal.WordWrap = True
    
    'Multi line textbox for the Description
    txtDescription.Enabled = False
    txtDescription.MultiLine = True
    txtDescription.WordWrap = True
        
    'So the excel file will not flash on screen when retrieving the data
    Application.ScreenUpdating = False
        
    'Open and sets reference to the workbook named InternationalTouristArrivalsWorldwideByRegion
    Dim wbCustomer As Workbook
    Set wbCustomer = Workbooks.Open(ThisWorkbook.Path & "\InternationalTouristArrivalsWorldwideByRegion.xlsx")
    
        'Retrieve all relevant information from the opened excel file
    With wbCustomer.Worksheets("Overview")
        txtSource.Value = Range("C10").Value
        txtConduct.Value = Range("C11").Value
        txtPeriod.Value = Range("C12").Value
        txtRegion.Value = Range("C13").Value
        txtPublished.Value = Range("C22").Value
        txtDate.Value = Range("C23").Value
        txtOriginal.Value = Range("C24").Value
        txtDescription.Value = Range("E9").Value
    End With
    
        'Close the workbook opened
    wbCustomer.Close
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub btnCancel_Click()
    
    Unload Me
    
End Sub

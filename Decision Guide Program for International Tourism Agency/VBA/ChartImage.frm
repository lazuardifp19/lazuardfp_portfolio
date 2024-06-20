VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartImage 
   Caption         =   "Chart Display"
   ClientHeight    =   11445
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15780
   OleObjectBlob   =   "ChartImage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ChartImageShower()
    
    btnCancel.Cancel = False
    
    Dim imageFile As String
    
    'Call different pictures for different options
    
    If ChartBuilder.lstOption.Value = "Stacked Area" Then
        imageFile = ThisWorkbook.Path & "\AnnualStackedArea.GIF"
        imgChart.Picture = LoadPicture(imageFile)
    End If
    
    If ChartBuilder.lstOption.Value = "Stacked Bar" Then
        imageFile = ThisWorkbook.Path & "\AnnualStackedBar.GIF"
        imgChart.Picture = LoadPicture(imageFile)
    End If
    
    If ChartBuilder.lstOption.Value = "Stacked Column" Then
        If ChartBuilder.optAnnual.Value = True Then
            imageFile = ThisWorkbook.Path & "\AnnualStackedColumn.GIF"
            imgChart.Picture = LoadPicture(imageFile)
        End If
    End If
    
    If ChartBuilder.lstOption.Value = "Clustered Bar" Then
        imageFile = ThisWorkbook.Path & "\StatisticalClusteredBar.GIF"
        imgChart.Picture = LoadPicture(imageFile)
    End If
    
    If ChartBuilder.lstOption.Value = "Clustered Column" Then
        imageFile = ThisWorkbook.Path & "\StatisticalClusteredColumn.GIF"
        imgChart.Picture = LoadPicture(imageFile)
    End If
    
    If ChartBuilder.lstOption.Value = "Stacked Column" Then
        If ChartBuilder.optStatistical.Value = True Then
            imageFile = ThisWorkbook.Path & "\StatisticalStackedColumn.GIF"
            imgChart.Picture = LoadPicture(imageFile)
        End If
    End If
    
    Me.Show
    
End Sub

Private Sub btnCancel_Click()

    Me.Hide
    btnCancel.Cancel = True
        
End Sub

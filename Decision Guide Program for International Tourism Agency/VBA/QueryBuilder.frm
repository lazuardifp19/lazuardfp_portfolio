VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueryBuilder 
   Caption         =   "Query Builder"
   ClientHeight    =   7935
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13785
   OleObjectBlob   =   "QueryBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim element As Integer
Dim i As Integer
Dim j As Integer
Dim counter As Integer

Dim regionArray(4, 1) As String
Dim subRegionArray(16, 1) As String
Dim countryArray(250, 2) As String
Dim gdpCountryArray(10000, 2) As String
Dim resultArray() As String

Public Sub CallQuery()
    
    Me.Show
    
End Sub

Private Sub lstCountry_Click()
    
    lstResult.Clear
    
    'Making the result array
    element = 0
    
    'Selected country option in the list box will be checked with the GDP Table
    For element = 0 To 10000
        
        'If the selected Country is found a match with the GDP table, proceed to dynamic array sizing
        If lstCountry.Value = gdpCountryArray(element, 1) Then
            GoTo arraySize
        End If
    Next element
    
arraySize:
    Dim index As Integer
    index = element
    
    i = 0
    j = 0
    
    'Checks for how many times the selected country name is displayed in the GDP Table
    For i = 0 To 49
        If lstCountry.Value = gdpCountryArray(index + i, 1) Then
            j = j + 1
        Else
            GoTo arrayResizing
        End If
    Next i

arrayResizing:
    
    'Subtraction for the indexing part, if j value alone, it will be a new country
    j = j - 1
    
    ReDim resultArray(0 To j, 0 To 2)
    
    Dim h As Integer
    h = UBound(resultArray)
    
    'Fetching the dates from the GDP table into the result array
    Dim a As Integer
    For a = 0 To j
        resultArray(a, 0) = gdpCountryArray(index + (h - a), 0)
    Next a
    
    'Fetching the names from the GDP table into the result array
    Dim b As Integer
    For b = 0 To j
        resultArray(b, 1) = gdpCountryArray(index + (h - b), 1)
    Next b
    
    'Fetching the value from the GDP table into the result array
    Dim c As Integer
    For c = 0 To j
        resultArray(c, 2) = gdpCountryArray(index + (h - c), 2)
    Next c
    
    'Getting the region and subregion
    element = 0
    
    'Comparing the selected country and comparing to a list of valid countries
    For element = 0 To 250
        If lstCountry.Value = countryArray(element, 2) Then
            GoTo regionIndexing
        End If
    Next element
    
regionIndexing:

    index = element
    
    a = 0
    b = 0
    
    For a = 0 To 4
        If regionArray(a, 0) = countryArray(index, 0) Then
            txtRegion.Value = regionArray(a, 1)
            GoTo subregion
        End If
    Next a
    
subregion:
    For b = 0 To 16
        If subRegionArray(b, 0) = countryArray(index, 1) Then
            txtSubRegion.Value = subRegionArray(b, 1)
            GoTo finalise
        End If
    Next b
    
finalise:
    
    lstResult.ColumnCount = 3
    lstResult.List() = resultArray
    lstResult.ListIndex = 0
    
    txtCountry = lstCountry.Value
    
End Sub

Private Sub lstResult_Click()
    
    'Update the GDP and Year Value on the text based on what result index is chosen
    txtGDP = resultArray(lstResult.ListIndex, 2)
    txtYear = resultArray(lstResult.ListIndex, 0)

End Sub

Private Sub UserForm_Initialize()
    
    lstCountry.Clear
    lstResult.Clear
    
    'Since the MSAccess is not changed for this project, the starting country is always Algeria
    Dim startingCountry As String
    startingCountry = "Algeria"
    
    'Make the text boxes not editable
    txtCountry.Enabled = False
    txtYear.Enabled = False
    txtGDP.Enabled = False
    txtRegion.Enabled = False
    txtSubRegion.Enabled = False
    
    'Call for the MSAccess data to be put into different arrays
    Call ACCDBConnectionArrays
    
    'Display all of the valid countries
    With lstCountry
        For element = 0 To UBound(countryArray)
            .AddItem countryArray(element, 2)
        Next element
    End With
    
    lstCountry.ListIndex = 0
    txtCountry.Value = startingCountry
    
    txtGDP = resultArray(0, 2)
    txtYear = resultArray(0, 0)
    
    'Making the result array
    element = 0
    
    For element = 0 To 10000
        If startingCountry = gdpCountryArray(element, 1) Then
            GoTo arraySize
        End If
    Next element
    
arraySize:
    Dim index As Integer
    index = element
    
    i = 0
    j = 0
    
    For i = 0 To 49
        If startingCountry = gdpCountryArray(index + i, 1) Then
            j = j + 1
        Else
            GoTo arrayResizing
        End If
    Next i

arrayResizing:
    
    j = j - 1
    
    ReDim resultArray(0 To j, 0 To 2)
    
    Dim h As Integer
    h = UBound(resultArray)
    
    Dim a As Integer
    For a = 0 To j
        resultArray(a, 0) = gdpCountryArray(index + (h - a), 0)
    Next a
    
    Dim b As Integer
    For b = 0 To j
        resultArray(b, 1) = gdpCountryArray(index + (h - b), 1)
    Next b
    
    Dim c As Integer
    For c = 0 To j
        resultArray(c, 2) = gdpCountryArray(index + (h - c), 2)
    Next c
    
    'Getting the region and subregion
    element = 0
    
    For element = 0 To 250
        If startingCountry = countryArray(element, 2) Then
            GoTo regionIndexing
        End If
    Next element
    
regionIndexing:

    index = element
    
    a = 0
    b = 0
    
    For a = 0 To 4
        If regionArray(a, 0) = countryArray(index, 0) Then
            txtRegion.Value = regionArray(a, 1)
            GoTo subregion
        End If
    Next a
    
subregion:
    For b = 0 To 16
        If subRegionArray(b, 0) = countryArray(index, 1) Then
            txtSubRegion.Value = subRegionArray(b, 1)
            GoTo finalise
        End If
    Next b
    
finalise:
    
    lstResult.ColumnCount = 3
    lstResult.List() = resultArray
    lstResult.ListIndex = 0
    
    txtCountry = startingCountry
    
End Sub

Sub ACCDBConnectionArrays()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim connectionRegionsADODB As New ADODB.Connection
    
    With connectionRegionsADODB
        
        'Path
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\Regions.accdb"
        
        'Provider
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        
        .Open
    End With
    
    'Regions record Set
    Dim rsRegionQuery As New Recordset
    rsRegionQuery.Open "SELECT * FROM Region", connectionRegionsADODB
    
    With rsRegionQuery
        element = 0
        Do Until .EOF
            regionArray(element, 0) = .Fields("RegionCode").Value
            regionArray(element, 1) = .Fields("RegionName").Value
            element = element + 1
            .MoveNext
        Loop
    End With
    
    'Sub-region record set
    Dim rsSubRegionQuery As New Recordset
    rsSubRegionQuery.Open "SELECT * FROM SubRegion", connectionRegionsADODB
    
    With rsSubRegionQuery
        element = 0
        Do Until .EOF
            subRegionArray(element, 0) = .Fields("SubregionCode").Value
            subRegionArray(element, 1) = .Fields("SubregionName").Value
            element = element + 1
            .MoveNext
        Loop
    End With
    
    'Country record set
    Dim rsCountryQuery As New Recordset
    Dim strCountry As String
    
    'SELECT DISTINCT to prevent any repetition of countries
    strCountry = "SELECT DISTINCT C.RegionCode, C.SubregionCode, C.CountryName FROM Country C, UN_GDP_ByCountry UN WHERE C.CountryName = UN.Country"
    rsCountryQuery.Open strCountry, connectionRegionsADODB
    
    With rsCountryQuery
        element = 0
        Do Until .EOF
            countryArray(element, 0) = .Fields("RegionCode").Value
            countryArray(element, 1) = .Fields("SubregionCode").Value
            countryArray(element, 2) = .Fields("CountryName").Value
            element = element + 1
            .MoveNext
        Loop
    End With
    
    'GDP per country
    Dim rsGDPCountryQuery As New Recordset
    Dim strGDP As String
    
    'Selecting UN GDP values from countries that are listed in the country database
    strGDP = "SELECT GD.Country, GD.Year, GD.Value FROM Country CN, UN_GDP_ByCountry GD WHERE GD.Country = CN.CountryName"
    rsGDPCountryQuery.Open strGDP, connectionRegionsADODB
    
    With rsGDPCountryQuery
        element = 0
        Do Until .EOF
            gdpCountryArray(element, 0) = .Fields("Year").Value
            gdpCountryArray(element, 1) = .Fields("Country").Value
            gdpCountryArray(element, 2) = .Fields("Value").Value
            element = element + 1
            .MoveNext
        Loop
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Private Sub btnCancel_Click()

    Unload Me

End Sub

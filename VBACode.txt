ThisWorkbook:

Option Explicit
Private Sub Workbook_Open()

Dim answer As String
Dim thisPath As String, datadirectory As String
Dim FSO As Object
' Clear old selections
Worksheets("Main").Select
 Range("A6:E6").Value = ""
 Range("A8:J8").Value = ""
 Call doCleaning
' Give a warning that we will change dot and comma as a decimal separator
  answer = MsgBox("We will change decimal to dot (.) and thousand separator will be empty! Is this OK or not?", vbYesNo + vbQuestion, "WARNING!")
        If answer = vbYes Then
         ' We set locale and then we continue
         Call SetLocale
           GoTo LetsContinue
         Else
           ' Stop execution
           Err.Raise 777, "Well you", "... wanted to stop? Just press end in this error message."
           Exit Sub
         End If

LetsContinue:
' Private Sub Workbook_Open()
' Here we do things what happens when workbook is opened
' First we check if Data Directory exists - if not, we assume data is not downloaded either
' First directory level variable
    thisPath = "C:\ExcelData"
    ' Second directory level variable
    datadirectory = "\AirportDataFiles\"
    ' check if the directory exist and if not, then we crate and download data
    Set FSO = CreateObject("Scripting.FileSystemObject")
        If Not FSO.FolderExists(thisPath & datadirectory) Then
       answer = MsgBox("Data directory: " & thisPath & datadirectory & " does not exists. Do you want to create it and download data?", vbYesNo + vbQuestion, "Data directory not exist!")
         If answer = vbYes Then
         ' We execute macro which create directory and downloads data
           Call loadData
         Else
           ' Stop execution
           Err.Raise 777, "Well you", "... wanted to stop? Just press end in this error message."
           Exit Sub
         End If
        End If
 ' Now ready to start with the Main sheet
    
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
 ' Here we do things what happens when workbook is closed
 ' Restore locale settings - use if needed. The close button will do this already
 ' Call RestoreLocale
End Sub



FormMain:

Option Explicit

Private Sub buttonClose_Click()
On Error GoTo Whoa
Application.EnableEvents = False

' Returns localization to original - you do not need this if your decimal separator is dot
Call RestoreLocale ' references at Parameters-page!
' Shuts down form
Unload Me
' Acticate the Main sheet
ThisWorkbook.Worksheets("Main").Activate

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue


End Sub

Private Sub cmdShowGoogleMaps_Click()
' Will show coordinates in the Google Map
Dim coordinates As String
' Read coordinates from the form field
 coordinates = ThisWorkbook.Sheets("Main").Range("B8").Value & "," & ThisWorkbook.Sheets("Main").Range("C8").Value
' Go to subs
 mcrURL (coordinates)

End Sub

Private Sub cmdShowHomeLink_Click()
Dim hyperLinkHome As String
' Pick up the hyperlink
hyperLinkHome = ThisWorkbook.Sheets("Main").Range("I8").Value
'Open the hyperlink
ThisWorkbook.FollowHyperlink (hyperLinkHome)

End Sub

Private Sub cmdShowWiki_Click()
Dim hyperLink As String
' Pick up the hyperlink
hyperLink = ThisWorkbook.Sheets("Main").Range("J8").Value
'Open the hyperlink
ThisWorkbook.FollowHyperlink (hyperLink)

End Sub

Private Sub CommandList_Click()

On Error GoTo Whoa

Application.EnableEvents = False

formMain.Caption = "Data analysis with Visual Basic (c) 2016 - Jari Hiltunen - This may take some time ... wait!"

' Let's clean old values if exists
Call doCleaning
' Inform user that this may take some time
Application.StatusBar = "Listing ... please, wait!"
' First we check which aircraft type was selected
Dim userSelection As String
userSelection = ThisWorkbook.Sheets("Main").Range("E6").Value
Select Case userSelection
    Case Is = "Helicopter"
    ' Now we go filters and data manipulation procedures
        Call filterHelicopter
    Case Is = "Private plane"
        Call filterPrivatePlane
    Case Is = "Passenger plane"
        Call filterPassangerPlane
    Case Is = "Cargo plane"
         Call filterCargoPlane
End Select
 ' Sort selections
 Call sortListed
Application.StatusBar = "List complete!"

ThisWorkbook.Worksheets("Main").Activate
' Update listBox items (otherwise you will see only 9 rows)
listAirportsBox.RowSource = "listAirports"
formMain.Caption = "Data analysis with Visual Basic (c) 2016 - Jari Hiltunen - All done!"

Me.Repaint

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue
End Sub
Private Sub listAirportsBox_Click()

On Error GoTo Whoa
Application.EnableEvents = False

' Updates BING/GoogleMaps based on selection
Dim newCoordinate As String
Dim airportName As String
Dim stringMyVal As String
Dim lngLastRow As Long
Dim strRowNoList As String
Dim bingLatitude As String
Dim bingLongitude As String
Dim targetDistance As String
Dim targetAirportType As String
Dim targetHomeLink As String
Dim targetWikiLink As String
Dim cell As Range
Dim url As String


airportName = listAirportsBox.Value
' Find selected airport row from the list
    stringMyVal = 1 'Value to search for, change as required.
    lngLastRow = Cells(Rows.Count, "A").End(xlUp).Row 'Search Column A
    For Each cell In Range("A11:A" & lngLastRow) 'Starting cell is A11
        If cell.Value = airportName Then
            If strRowNoList = "" Then
                strRowNoList = strRowNoList & cell.Row
            Else
                strRowNoList = strRowNoList & ", " & cell.Row
            End If
        End If
    Next cell
'Construct coordinates from the selected row (this is for Google which does not work
newCoordinate = Worksheets("Main").Range("B" & strRowNoList).Value & "," & Worksheets("Main").Range("C" & strRowNoList).Value
' Contruct address for BingBong
bingLatitude = Worksheets("Main").Range("B" & strRowNoList).Value
bingLongitude = Worksheets("Main").Range("C" & strRowNoList).Value
' Pick up distance information
targetDistance = Worksheets("Main").Range("D" & strRowNoList).Value
' Pick up airport type information
targetAirportType = Worksheets("Main").Range("F" & strRowNoList).Value
' Pick up homelink type information
targetHomeLink = "=Hyperlink(" & Chr(34) & Worksheets("Main").Range("I" & strRowNoList).Value & Chr(34) & ")"
' Pick up wikilink type information
targetWikiLink = "=Hyperlink(" & Chr(34) & Worksheets("Main").Range("J" & strRowNoList).Value & Chr(34) & ")"

' Update Google Maps with new addresses ?force=canvas/ ?force=lite.
      ' url = "http://www.google.com/maps/place/" & newCoordinate
      ' Does not work in Windows 10 and Internet Explorer 11!
' Use Bing maps instead
' This is not so nice, but there is no other ways to my unserstanding to fix Excel OLE component (uses always IE11)
       url = "http://www.bing.com/maps/default.aspx?v=2&cp=" & bingLatitude & "~" & bingLongitude & "&lvl=12"
   Call OpenURLinForms(url)
' Update main-sheet about selection
' Airport name to A8
Worksheets("Main").Range("A8").Value = airportName
' Latitude to B8
Worksheets("Main").Range("B8").Value = bingLatitude
' Latitude to C8
Worksheets("Main").Range("C8").Value = bingLongitude
' Distance to D8
Worksheets("Main").Range("D8").Value = targetDistance
' Airport type to F8
Worksheets("Main").Range("F8").Value = targetAirportType
' Homelink to I8
Worksheets("Main").Range("I8").Value = targetHomeLink
' Wikilink to J8
Worksheets("Main").Range("J8").Value = targetWikiLink
' Now this information can be parsed to the web browser or forms if needed

' Enable buttons possible
' The Wikibutton
If (ThisWorkbook.Sheets("Main").Range("J8").Value <> "no information") Then
  cmdShowWiki.Enabled = True
  Else 'Means that link is missing
   cmdShowWiki.Enabled = False
End If

' The homelink button
If (ThisWorkbook.Sheets("Main").Range("I8").Value <> "no information") Then
  cmdShowHomeLink.Enabled = True
  Else
  cmdShowHomeLink.Enabled = False
End If

'Enable Show in Google Maps button
cmdShowGoogleMaps.Enabled = True

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub
Private Sub tboxDestinationAddress_AfterUpdate()
' We execute this macro after we loose focus on the field
On Error GoTo Whoa
Application.EnableEvents = False

Dim destinationAddress As String
Dim coordinates As String, longitude As String, latitude As String
Dim coordinatesArray() As String, url As String

' We do nothing if address is less than 5 characters!
If Len(tboxDestinationAddress.Value) > 5 Then
' Read value from the field
  destinationAddress = tboxDestinationAddress.Value
 Worksheets("Main").Activate
 ' Solve coordinates
 coordinates = GoogleGeocode(destinationAddress)
 tboxAddressUpdate.Value = coordinates
 ' Write coordinates to the sheet
 Worksheets("Main").Activate
  Range("A6").Value = destinationAddress
  ' This is not used but left as an example.
  ' We have to manipulate coordinates so that we separate longitude and latitude and
  ' then replace . to , because imported coordinates are presented in filand locale!
  ' you shall change . to , script if other locale used!
  If (Len(coordinates) > 12) Then ' Greater than word "Not found!"
  ' Enable Distance field
    tboxMaxDistanceKM.Enabled = True
    coordinatesArray() = Split(coordinates, ",")
  ' Array 0 is first block before separator (,) and now we replace . to , (Finland locale)
  ' latitude = Replace(coordinatesArray(0), ".", ",")
  latitude = coordinatesArray(0)
   Range("B6").Value = latitude
  ' longitude = Replace(coordinatesArray(1), ".", ",")
  longitude = coordinatesArray(1)
   Range("C6").Value = longitude
  ' Let's update the Google Maps ?force=canvas/
  ' Google DOES NOT work with Windows 10 and IE 11!
  '    url = "https://www.google.com/maps/place/" & coordinates
  ' Therefore we use BING maps
    url = "http://www.bing.com/maps/default.aspx?v=2&cp=" & latitude & "~" & longitude & "&lvl=12"
    Call OpenURLinForms(url)
 End If
  ' Autofit
  Worksheets("Main").Range("A:J").Columns.AutoFit
  Else
   Exit Sub
End If



LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub

Private Sub tboxMaxDistanceKM_Change()
' Error handling
On Error GoTo Whoa

Application.EnableEvents = False

Dim maxDistance As Integer
Dim inputValue As Variant
' Check validity of the number
   inputValue = tboxMaxDistanceKM.Value
   If Not IsNumeric(inputValue) Then
    'is not a number
    MsgBox ("Please, numbers only!")
  Else
    ' Fine, let's put value in the sheet
  maxDistance = inputValue
  Worksheets("Main").Activate
  Range("D6").Value = maxDistance
  ' Autofit
  Worksheets("Main").Range("A:J").Columns.AutoFit
  End If
  
  ' Enable the List airports button after distance is inputted and address resolved properly!
  If (Len(comboAircraftSelect.Value > 0)) & (Val(tboxAddressUpdate.Value > 0)) Then
       CommandList.Enabled = True
  End If

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue
 
End Sub

Private Sub UpdateData_Click()
' Call sub which load data
Call loadData
' Refresh connections (reloads data into sheets)
    ActiveWorkbook.RefreshAll
 ' Inform user (data update status messages also in the bottom field)
    MsgBox ("Data downloaded and connections to the data refressed!")
End Sub

' This MUST be first if you would like to populate names from name ranges!
Private Sub UserForm_Initialize()
' Error handling
On Error GoTo Whoa

Application.EnableEvents = False
' Let's hide buttons so far they are ready to be enabled
CommandList.Enabled = False: cmdShowWiki.Enabled = False: cmdShowHomeLink.Enabled = False: cmdShowGoogleMaps.Enabled = False:

' This is for Aircraft selection!
Dim aircraftType As Range
Dim ws As Worksheet
Dim rngName As Range
Set ws = Worksheets("Parameters")
' Name range Aircrafts contains =OFFSET(Parameters!$D$2;0;0;COUNTA(Parameters!$D:$D)-1;1)
For Each aircraftType In ws.Range("Aircrafts")
  Me.comboAircraftSelect.AddItem aircraftType.Value
Next aircraftType
' Clear old selections
Worksheets("Main").Select
 Range("A6:E6").Value = ""
 Range("A8:J8").Value = ""
 
' Pick up airport name range for the form
listAirportsBox.RowSource = "listAirports"
   
LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub

Private Sub comboAircraftSelect_Change()
Dim AircraftSelection As String
  ' Read selected value
  AircraftSelection = comboAircraftSelect.Value
  ' Put value in the sheet
  Worksheets("Main").Activate
  Range("E6").Value = AircraftSelection
  ' Autofit
  Worksheets("Main").Range("A:J").Columns.AutoFit
  ' Disable field field
  tboxMaxDistanceKM.Enabled = False
  
End Sub

Sub OpenURL(urlName)
' If you like to use some other browser, change object reference!
' This is if you would like to use from command
 Dim ie As Object
  Set ie = CreateObject("InternetExplorer.Application")
  With ie
  .Navigate urlName
  .Visible = True
  End With
 
End Sub

Sub mcrURL(coordinates As String)
' Parses proper URL from coordinates
Dim url As String
' If you want to use satellite map type, then append /data=!3m1!1e3
' If you want terrain view of the map, then append /data=!3m1!4b1
  url = "https://www.google.com/maps/place/" & coordinates
  Call OpenURL(url)
End Sub

Sub OpenURLinForms(urlName)
' If you like to use some other browser, change object reference!
' This is related to userForms
  Me.showGoogleWeb.Navigate urlName
End Sub



Module 1:

Option Explicit
Sub main()
' This is referenced from sheet Main and from button START
On Error GoTo Whoa
Application.EnableEvents = False
' Clean old things away
Call doCleaning
' Change locale so that decimal is dot and no thousand separators
Call SetLocale
' Bring up user form (this could be automated so that kicks in when Excel starts
formMain.Show vbModeless

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub

Sub loadData()
' Create directory for Data and download needed files to txt-format.
On Error GoTo Whoa
Application.EnableEvents = False

Dim myURL As String
' Pass URL and target file to function, which does the job
' You may want to add more features and link data to existing forms ... here are ready all commands for download!
Application.StatusBar = "Begin downloading Airports data"
' First parameter is the URL and file to be downloaded, second is file to be saved!
' Because Excel can not directly handle UTF-8 characters, csv need to go via txt
downLoadFromURL "http://ourairports.com/data/airports.csv", "airports.txt"
'Application.StatusBar = "Begin downloading Airport-frequencies data"
'downLoadFromURL "http://ourairports.com/data/airport-frequencies.csv", "afrq.txt"
'Application.StatusBar = "Begin downloading Runways data"
'downLoadFromURL "http://ourairports.com/data/runways.csv", "runways.txt"
'Application.StatusBar = "Begin downloading Countries data"
'downLoadFromURL "http://ourairports.com/data/countries.csv", "countries.txt"
'Application.StatusBar = "Begin downloading Regions data"
'downLoadFromURL "http://ourairports.com/data/regions.csv", "regions.txt"
'Application.StatusBar = "Begin downloading Worldwide radio navigation aids"
'downLoadFromURL "http://ourairports.com/data/navaids.csv", "navids.txt"
'Application.StatusBar = "All downloads complete!"
 
' If not already loaded, next import data to the sheets
' Application.StatusBar = "Loading data to the Airport sheet ... wait"
 ' Application.Run ("LoadAirportData")
' Normally it is enough to update connections
ActiveWorkbook.RefreshAll
  Call listFilesToParameters
Application.StatusBar = "All completed!"

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub
Function downLoadFromURL(myURL As String, myFile As String)
' Takes in URL with absolute path and myFile as target filename
Dim thisPath As String, ext As String, myName As String, datadirectory As String
Dim WinHttpReq As Object, FSO As Object, oStream As Object
' Application.EnableEvents = False
' Define path where we shall save datafiles
' For this project I have to save these files into C:, because Excel do not support RELATIVE PATHS!
' Meaning: if I add files downloaded by this script to Excel Data Model, they will point to absolute paths
' DO NOT begin directory name with blank! It is ILLEGAL for OneDrive!
    ' First directory level variable
    thisPath = "C:\ExcelData"
    ' Second directory level variable
    datadirectory = "\AirportDataFiles\"
    ' check if the directory exist and if not, create it
    Set FSO = CreateObject("Scripting.FileSystemObject")
        If Not FSO.FolderExists(thisPath & datadirectory) Then
        ' First level directory
        FSO.CreateFolder (thisPath)
        ' Second level directory
        FSO.CreateFolder (thisPath & datadirectory)
    End If
' Begin grabbing
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False, "username", "password"
WinHttpReq.send
myURL = WinHttpReq.responseBody
' If status = ok, then download
If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    ' Path where files will be stored
    oStream.SaveToFile thisPath & "\" & datadirectory & "\" & myFile, 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If

End Function

Sub listFilesToParameters()
' List files downloaded to the parameters page
' This is not neccessary to run, but for may help finding problems
Worksheets("Parameters").Activate
Range("A1").Value = "Filename"
Range("B1").Value = "Path & Filename"
Dim varDirectory As Variant
Dim flag As Boolean
Dim i As Integer
Dim strDirectory As String, myName As String, datadirectory As String
    myName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".") - 1))
    ' Obs! Here is static "datafiles" which should be changed in all subs (not passed as parameter!)
    datadirectory = "AirportDataFiles"
    ' See static setting to C:\ExcelData
    strDirectory = "C:\ExcelData" & "\" & datadirectory & "\"
    ' Iterator
    i = 1
    flag = True
    varDirectory = Dir(strDirectory, vbNormal)
    While flag = True
    If varDirectory = "" Then
        flag = False
    Else
        Cells(i + 1, 1) = varDirectory
        Cells(i + 1, 2) = strDirectory + varDirectory
        'returns the next file or directory in the path
        varDirectory = Dir
        i = i + 1
    End If
Wend
' Fit the columns
Worksheets("Parameters").Range("A:B").Columns.AutoFit

End Sub
Public Function GetDistanceCoord(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double, ByVal unit As String) As Double
' Calculate distance based on given latitude and latitude. Last variable is either K = Kilometers, N = Miles
    Dim theta As Double: theta = lon1 - lon2
    Dim dist As Double: dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta))
    dist = WorksheetFunction.Acos(dist)
    dist = rad2deg(dist)
    dist = dist * 60 * 1.1515
    ' For Kilometers
    If unit = "K" Then
        dist = dist * 1.609344
    ElseIf unit = "N" Then
        dist = dist * 0.8684
    End If
    GetDistanceCoord = dist
End Function
Function deg2rad(ByVal deg As Double) As Double
 ' Degrees to radians
    deg2rad = (deg * WorksheetFunction.Pi / 180#)
End Function
Function rad2deg(ByVal rad As Double) As Double
 ' Radians to degrees
    rad2deg = rad / WorksheetFunction.Pi * 180#
End Function


Module 2:

Option Explicit
Function GoogleGeocode(address As String) As String
' Function returns longitude and latitude (code modified from policeanalyst.com code)
  Dim strAddress As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String

  strAddress = URLEncode(address)

  'Assemble the query string
  strQuery = "http://maps.googleapis.com/maps/api/geocode/xml?"
  strQuery = strQuery & "address=" & strAddress
  strQuery = strQuery & "&sensor=false"

  ' Define XML and HTTP components - fixed for the version 6 and for new version MSXML!
  ' Remember to add from Tools - References Microsoft XML v6!
  Dim googleResult As Object
  Dim googleService As Object
  Set googleResult = CreateObject("Msxml2.DOMDocument.6.0")
  Dim ie As MSXML2.XMLHTTP60
  Set ie = CreateObject("MSXML2.XMLHTTP.6.0")
  Set googleService = CreateObject("MSXML2.XMLHTTP")
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode
  
  'create HTTP request to query URL - make sure to have
  'that last "False" there for synchronous operation
  googleService.Open "GET", strQuery, False
  googleService.send
  googleResult.LoadXML (googleService.responseText)

  Set oNodes = googleResult.getElementsByTagName("geometry")

  If oNodes.Length = 1 Then
    For Each oNode In oNodes
      strLatitude = oNode.ChildNodes(0).ChildNodes(0).Text
      strLongitude = oNode.ChildNodes(0).ChildNodes(1).Text
      GoogleGeocode = strLatitude & "," & strLongitude
    Next oNode
  Else
    GoogleGeocode = "Not Found!"
  End If
End Function
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
' Here we encode the URL to be passed for the coordinate
' So far this works with Google, but if this stops working, BING MAY be used with some modifications
  Dim StringLen As Long: StringLen = Len(StringVal)
  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String
    
    If SpaceAsPlus Then Space = "+" Else Space = "%20"
    
    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      
      Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function


Module 3:

Option Explicit
Sub doCleaning()
' This sub will do cleaning to calculated areas
ThisWorkbook.Worksheets("Main").Activate
' Begin from A11
 ActiveSheet.Range("A11", ActiveSheet.Range("J11").End(xlDown)).Clear
 
End Sub
Sub SetLocale()
' This module handles annoying locales issue with decimal separator etc
' Define separators and apply.
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = "-"
    Application.UseSystemSeparators = False
    ' Change column decimals - not really needed for other countries
    Sheets("Main").Range("D:D").NumberFormat = "0.0"

   
End Sub
Sub RestoreLocale()
 ' Restore local settings.
 Dim restoreDecimal As String, restoreThousands As String
 restoreDecimal = Worksheets("Parameters").Range("H3").Value
 restoreThousands = Worksheets("Parameters").Range("I3").Value
 ' Define separators and apply.
    Application.DecimalSeparator = restoreDecimal
    Application.ThousandsSeparator = restoreThousands
    Application.UseSystemSeparators = False
End Sub
   
Sub sortListed()
' Sorts listed airports in the sheet by distance (column D)
' For some reason more intelligent way to sort fails. This is quick and dirty.
Range("A10").Select
    Range(selection, selection.End(xlToRight)).Select
    Range(selection, selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Main").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Main").Sort.SortFields.Add Key:=Range("D11:D65535") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Main").Sort
        ' Most likely newer reach this much
        .SetRange Range("A10:J65535")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Module 6:

Option Explicit
Sub filterHelicopter()
' Filters data for Helicopters suitable airports
Application.ScreenUpdating = False
    Sheets("Airportdata").Select
    ActiveSheet.ListObjects("airports").Range.AutoFilter Field:=2, Criteria1:= _
       "=*heliport*", Operator:=xlOr, Criteria2:="=*small_airport*"
' Add distance information to the sheet
Call calculateDistance
Call listAirportsinRange
Application.ScreenUpdating = True
End Sub
Sub filterPrivatePlane()
' Filters data for Private plane suitable airports
Application.ScreenUpdating = False
    Sheets("Airportdata").Select
       ActiveSheet.ListObjects("airports").Range.AutoFilter Field:=2, Criteria1:= _
       "=*seaplane_base*", Operator:=xlOr, Criteria2:="=*medium_airport*"
' Add distance information to the sheet
Call calculateDistance
Call listAirportsinRange
Application.ScreenUpdating = True
End Sub
Sub filterPassangerPlane()
' Filters data for Passanger plane suitable airports
Application.ScreenUpdating = False
    Sheets("Airportdata").Select
      ActiveSheet.ListObjects("airports").Range.AutoFilter Field:=2, Criteria1:= _
       "=*large_airport*", Operator:=xlOr, Criteria2:="=*large_airport*"
' Add distance information to the sheet
Call calculateDistance
Call listAirportsinRange
Application.ScreenUpdating = True
End Sub
Sub filterCargoPlane()
' Filters data for Cargo plane suitable airports
Application.ScreenUpdating = False
    Sheets("Airportdata").Select
        ActiveSheet.ListObjects("airports").Range.AutoFilter Field:=2, Criteria1:= _
       "=*medium_airport*", Operator:=xlOr, Criteria2:="=*large_airport*"
' Add distance information to the sheet
Call calculateDistance
Call listAirportsinRange
Application.ScreenUpdating = True
End Sub
Sub calculateDistance()
Dim sht As Worksheet
Dim LastRow As Long
Dim Rows As Long
Dim oLatitude As Double, oLongitude As Double
Dim res As String
Dim LC As Long, i As Long
Dim rng As Range


On Error GoTo Whoa
Application.EnableEvents = False

Application.ScreenUpdating = False
Sheets("Airportdata").Select
' This sub calculates distances for filtered airports
' Calculation is FLAT WORLD calculation, because Google-calculation will take too much time
' First we delete previous cDistance column if exist
res = "cDistance"
LC = Cells(1, Columns.Count).End(xlToLeft).Column
Set rng = Range(Cells(1, 2), Cells(1, LC))
    For i = LC To 2 Step -1  'continue loop from last column to Column B (chg is needed)
        If Cells(1, i).Value = res Then
        Cells(1, i).EntireColumn.Delete
        End If
    Next i
' Add (perhaps again) new column
ActiveSheet.Columns(6).Insert
' Rename column
Range("F1") = "cDistance"
' Calculate first and last rows
Set sht = ThisWorkbook.Worksheets("Airportdata")
'Ctrl + Shift + End
  LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
  Rows = LastRow - 1
ActiveSheet.Columns(4).Select
' Lets calculate
' Pick up values from Main page to variables
oLatitude = Val(ThisWorkbook.Sheets("Main").Range("origLatitude"))
oLongitude = Val(ThisWorkbook.Sheets("Main").Range("origLongitude"))
For i = 2 To Rows
'' Here Cell 6 = calcualated distance, cell 4 = latitude of the airport, cell 5 = longitude of the airport, K = kilometers
'' Values are Double-values (Val) and DO NOT use CDbl in Finland! CDBl uses SYSTEM LOCALE for decimal pointers!
Cells(i, 6).Value = GetDistanceCoord(oLatitude, oLongitude, Val(Cells(i, 4).Value), Val(Cells(i, 5).Value), "K")
Next i
Application.ScreenUpdating = True

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub
Sub listAirportsinRange()
Dim cl As Range, rng As Range
Dim x As Long

On Error GoTo Whoa
Application.EnableEvents = False
Application.ScreenUpdating = False
' Lists airports within selected range
' Begin row at Main page x = iterator
  x = 11
 ThisWorkbook.Worksheets("Airportdata").Activate
 ' Set range of the filtered range
   Set rng = Range("A2", Range("A2").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)
  ' Repeat for each filtered row
    For Each cl In rng.SpecialCells(xlCellTypeVisible)
    'cl = row which is filtered
    ' Then we list all Airports which suites for the criteria
    ' Obs! Column 7 is new created column with distances!
     If (Val(cl.Offset(0, 5).Value) <= Val(ThisWorkbook.Sheets("Main").Range("distance"))) Then
      ' Set Distance to Main sheet to D
     ThisWorkbook.Sheets("Main").Cells(x, 4).Value = Val(cl.Offset(0, 5))
     ' Set Name to Main sheet to A
     ThisWorkbook.Sheets("Main").Cells(x, 1).Value = (cl.Offset(0, 2))
     ' Set country to H
     ThisWorkbook.Sheets("Main").Cells(x, 8).Value = (cl.Offset(0, 7))
    ' Set Continent to Main sheet to G
     ThisWorkbook.Sheets("Main").Cells(x, 7).Value = (cl.Offset(0, 8))
     ' Set Municipality to Main sheet to E
     ThisWorkbook.Sheets("Main").Cells(x, 5).Value = (cl.Offset(0, 9))
    ' Set Airport type to Main sheet to F
    ThisWorkbook.Sheets("Main").Cells(x, 6).Value = (cl.Offset(0, 1))
    ' Set Airport latitude to Main sheet to B
    ThisWorkbook.Sheets("Main").Cells(x, 2).Value = Val(cl.Offset(0, 3))
    ' Set Airport longitude to Main sheet to C
    ThisWorkbook.Sheets("Main").Cells(x, 3).Value = Val(cl.Offset(0, 4))
    ' Set Wikipedia to I
     ThisWorkbook.Sheets("Main").Cells(x, 9).Value = (cl.Offset(0, 10))
    ' If value is empty, we write no information
    If (ThisWorkbook.Sheets("Main").Cells(x, 9).Value) = "" Then
      ThisWorkbook.Sheets("Main").Cells(x, 9).Value = "no information"
    End If
    ' Set Homelink to J
       ThisWorkbook.Sheets("Main").Cells(x, 10).Value = (cl.Offset(0, 11))
        ' If value is empty, we write no information
     If (ThisWorkbook.Sheets("Main").Cells(x, 10).Value) = "" Then
       ThisWorkbook.Sheets("Main").Cells(x, 10).Value = "no information"
     End If
    ' Iterate Main page downwards
      x = x + 1
   End If
  Next cl

 ' Fit the columns
Worksheets("Main").Range("A:J").Columns.AutoFit
' Change column decimals - not really needed for other countries
Sheets("Main").Range("D:D").NumberFormat = "0.0"

Application.ScreenUpdating = True

' Select main
ThisWorkbook.Worksheets("Main").Activate

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue

End Sub

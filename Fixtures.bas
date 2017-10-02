Attribute VB_Name = "Fixtures"
Option Explicit

Public ColFixtures As Collection

Public arrCompetitions() As String
Public arrDates() As String


Public dicLeaguePositions As Object

'Mapping table. The names used in the fixtures and league tables are different so need to be mapped consistently
Public dicTeamNames As Object
Public arrTeamFilter() As String

'==============================
'Name:MakeReport
'Purpose: Main sub. Gets the data from worksheets and loads it into variables
'         when the setup is complete it loads the form for creating reports.
'==============================
Private Sub MakeReport()

Set ColFixtures = New Collection
Set dicLeaguePositions = CreateObject("Scripting.Dictionary")
Set dicTeamNames = CreateObject("Scripting.Dictionary")
Dim wsPremierLeagueTable As Worksheet
Dim wsChampionshipTable As Worksheet
Dim wsLeagueOneTable As Worksheet
Dim wsLeagueTwoTable As Worksheet
Dim wsMapping As Worksheet
Dim wsFilter As Worksheet
Dim booContinueMacro As Boolean

Set wsFilter = Application.ThisWorkbook.Sheets("Filter")
Set wsMapping = Application.ThisWorkbook.Sheets("Mapping")
Set wsPremierLeagueTable = Application.ThisWorkbook.Sheets("Premier_League_Table")
Set wsChampionshipTable = Application.ThisWorkbook.Sheets("Championship_Table")
Set wsLeagueOneTable = Application.ThisWorkbook.Sheets("League_One_Table")
Set wsLeagueTwoTable = Application.ThisWorkbook.Sheets("League_Two_Table")

booContinueMacro = CreateTeamFilter(arrTeamFilter, wsFilter)
If booContinueMacro = False Then
    MsgBox ("Error occurred when creating filter. Macro stopping")
    Exit Sub
End If

booContinueMacro = MapTeamNames(dicTeamNames, wsMapping)

If booContinueMacro = False Then
    MsgBox ("Error occurred when mapping team names to match names used in fixtures with those used in league tables. (These are sometimes shortened in the league tables). Macro stopping")
    Exit Sub
End If


'Premier League
booContinueMacro = GetLeaguePositions(dicLeaguePositions, wsPremierLeagueTable, dicTeamNames)

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting league positions for teams in the Premier League. Macro stopping.")
    Exit Sub
End If

'Championship
booContinueMacro = GetLeaguePositions(dicLeaguePositions, wsChampionshipTable, dicTeamNames)

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting league positions for teams in the Championship. Macro stopping.")
    Exit Sub
End If

'League One
booContinueMacro = GetLeaguePositions(dicLeaguePositions, wsLeagueOneTable, dicTeamNames)

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting league positions for teams in League One. Macro stopping.")
    Exit Sub
End If

booContinueMacro = GetLeaguePositions(dicLeaguePositions, wsLeagueTwoTable, dicTeamNames)

'League Two
If booContinueMacro = False Then
    MsgBox ("Error occurred while getting league positions for teams in League Two. Macro stopping.")
    Exit Sub
End If



booContinueMacro = GetFixtureList()

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting the fixture list and transforming it into a data table. Macro stopping.")
    Exit Sub
End If


booContinueMacro = GetDistinctCompetitions(arrCompetitions)

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting the list of distinct competitions. Macro stopping.")
    Exit Sub
End If

booContinueMacro = GetDistinctDates(arrDates)

If booContinueMacro = False Then
    MsgBox ("Error occurred while getting the list of distinct match dates. Macro stopping.")
    Exit Sub
End If

frmFixtures.Show

End Sub

'==============================
'Name:MapTeamNames
'Purpose: To map team names for fixtures to team names for leagues (as these do not always match)
'         Params are the dictionary that the data will be stored in and the mapping worksheet
'==============================

Private Function MapTeamNames(ByRef dicTeamNames, ByVal wsSheet As Worksheet) As Boolean
On Error GoTo errHandler
Dim i As Long

For i = 1 To wsSheet.UsedRange.Rows.Count
    dicTeamNames.Add wsSheet.Cells(i, 1).value, wsSheet.Cells(i, 2).value
Next i

MapTeamNames = True
Exit Function

errHandler:
    MapTeamNames = False
End Function

'==============================
'Name: GetLeaguePositions
'Purpose: Gets the league position data from the worksheet it is stored in and populates
'         the dicLeaguePositions dictionary for future use.
'==============================

Private Function GetLeaguePositions(ByRef dicLeaguePositions, ByVal wsSheet As Worksheet, ByRef dicTeamNames) As Boolean
On Error GoTo errHandler
Dim i As Long
Dim strTeamName As String

For i = 3 To wsSheet.UsedRange.Rows.Count 'not interested in headers
    strTeamName = wsSheet.Cells(i, 2)
    dicLeaguePositions.Add dicTeamNames(strTeamName), wsSheet.Cells(i, 1).value
Next i

GetLeaguePositions = True
Exit Function

errHandler:
    GetLeaguePositions = False
End Function

'==============================
'Name:FormatReport
'Purpose: Formats ranges according to a level 1, 2, or 3. 1 = header 1,
'           2 = header 2, 3 =body text
'==============================

Private Sub FormatReport(ByVal intLevel As Integer, ByVal wsRange As Range)
Dim intFontSize As Integer
Dim booBold As Boolean
Dim booItalic As Boolean

If intLevel = 1 Then
    intFontSize = 15
    booBold = True
    booItalic = True
ElseIf intLevel = 2 Then
    intFontSize = 13
    booBold = True
    booItalic = False
Else: intFontSize = 12
    booBold = False
    booItalic = False
End If

With wsRange
    .Font.Size = intFontSize
    .Font.Name = "Verdana"
    .Font.Bold = booBold
    .Font.Italic = booItalic
End With


End Sub

'==============================
'Name:AddFixtures
'Purpose: Adds a new set of fixtures to the output report depending on data received
'         from the form.
'==============================

Public Sub AddFixtures( _
    ByVal strCompetition As String, _
    ByVal strDate As String, _
    ByVal booFiltered As Boolean, _
    ByVal booIncludeDate As Boolean)
    

Dim booIncludeRow As Boolean
Dim wsFixturesReport As Worksheet
Set wsFixturesReport = Application.ThisWorkbook.Sheets("Fixtures_Report")
Dim lonNextRow As Long
Dim strHomeTeamLeaguePosition As String
Dim strAwayTeamLeaguePosition As String

booIncludeRow = True
lonNextRow = wsFixturesReport.UsedRange.Rows.Count

If lonNextRow <> 1 Then
    lonNextRow = lonNextRow + 1
End If

If booIncludeDate = True Then
    wsFixturesReport.Cells(lonNextRow, 1) = strDate
    FormatReport 1, wsFixturesReport.Range(wsFixturesReport.Cells(lonNextRow, 1), wsFixturesReport.Cells(lonNextRow, 1))
    lonNextRow = lonNextRow + 1
End If

wsFixturesReport.Cells(lonNextRow, 1) = strCompetition
FormatReport 2, wsFixturesReport.Range(wsFixturesReport.Cells(lonNextRow, 1), wsFixturesReport.Cells(lonNextRow, 1))
lonNextRow = lonNextRow + 1

Dim fixture As cFixture

For Each fixture In ColFixtures

    If fixture.Competition = strCompetition Then

        If fixture.MatchDate = strDate Then
            If booFiltered = False Then
                booIncludeRow = True
            Else

                booIncludeRow = CheckFilters(fixture.HomeTeam, fixture.AwayTeam)
            End If
                
            If booIncludeRow = True Then
            

                strHomeTeamLeaguePosition = dicLeaguePositions(fixture.HomeTeam)
                
                If strHomeTeamLeaguePosition <> vbNullString Then
                strHomeTeamLeaguePosition = " (" & strHomeTeamLeaguePosition & ")"
                End If
                
                strAwayTeamLeaguePosition = dicLeaguePositions(fixture.AwayTeam)
                
                If strAwayTeamLeaguePosition <> vbNullString Then
                strAwayTeamLeaguePosition = " (" & strAwayTeamLeaguePosition & ")"
                End If
                
                wsFixturesReport.Cells(lonNextRow, 1).value = fixture.HomeTeam & strHomeTeamLeaguePosition
                wsFixturesReport.Cells(lonNextRow, 2) = "v"
                wsFixturesReport.Cells(lonNextRow, 3).value = fixture.AwayTeam & strAwayTeamLeaguePosition
                wsFixturesReport.Cells(lonNextRow, 4).value = fixture.KickOff
                
                FormatReport 3, wsFixturesReport.Range(wsFixturesReport.Cells(lonNextRow, 1), wsFixturesReport.Cells(lonNextRow, 4))
            
                lonNextRow = lonNextRow + 1
            End If
        End If
    End If
    
Next fixture

wsFixturesReport.Range("A:B:C:D").EntireColumn.AutoFit
With wsFixturesReport.PageSetup
    .Zoom = False
    .FitToPagesWide = 1
End With
wsFixturesReport.Select
End Sub

'==============================
'Name:CheckFilters
'Purpose: Checks to see if the home team or away team is one that should be filtered on.
'         E.g. if the filter is on London teams then does either team fit in that category.
'==============================

Private Function CheckFilters(ByVal strHomeTeam As String, ByVal strAwayTeam As String) As Boolean
Dim i As Long

CheckFilters = False
For i = 1 To UBound(arrTeamFilter)
If arrTeamFilter(i) = strHomeTeam Or arrTeamFilter(i) = strAwayTeam Then
    CheckFilters = True
End If
Next i

End Function

'==============================
'Name:GetDistinctDates
'Purpose:  Get a distinct list of dates from the fixtures list to populate the dates drop down
'          in the form.
'==============================

Private Function GetDistinctDates(ByRef arrDates() As String) As Boolean
On Error GoTo errHandler
Dim dictDates As Object
Set dictDates = CreateObject("Scripting.Dictionary")

Dim key As Variant
Dim lonRecCount As Long
lonRecCount = 1

Dim fixture As cFixture

For Each fixture In ColFixtures
    dictDates(fixture.MatchDate) = 1
Next fixture

For Each key In dictDates.Keys
    
    ReDim Preserve arrDates(lonRecCount)
    arrDates(lonRecCount) = key
    lonRecCount = lonRecCount + 1
Next key

Set dictDates = Nothing

GetDistinctDates = True
Exit Function

errHandler:
    GetDistinctDates = False

End Function

'==============================
'Name:      GetDistinctCompetitions
'Purpose:   To get a distinct list of competitions from the fixtures list
'           to populate the form's drop down list.
'==============================

Private Function GetDistinctCompetitions(ByRef arrCompetitions() As String) As Boolean
On Error GoTo errHandler
Dim dictCompetitions As Object
Set dictCompetitions = CreateObject("Scripting.Dictionary")

Dim key As Variant
Dim lonRecCount As Long
lonRecCount = 1
Dim fixture As cFixture

For Each fixture In ColFixtures
    dictCompetitions(fixture.Competition) = 1
Next fixture

For Each key In dictCompetitions.Keys
    
    ReDim Preserve arrCompetitions(lonRecCount)
    arrCompetitions(lonRecCount) = key
    lonRecCount = lonRecCount + 1
    
Next key

Set dictCompetitions = Nothing

GetDistinctCompetitions = True
Exit Function

errHandler:
    GetDistinctCompetitions = False
End Function

'==============================
'Name:    GetFixtureList
'Purpose: To get the fixture list and transform it into a useable format.
'         Fixtures are recorded as with Date as header 1, Competition as header 2,
'         and Fixtures as the body. This is the intended format but since you can't filter this
'         to show only the competitions or teams you are interested in this needs to be tranformed
'         into a standard tablular fomat.
'==============================

Private Function GetFixtureList() As Boolean

Dim wsFixtures As Worksheet
Set wsFixtures = Application.ThisWorkbook.Sheets("Fixtures")
Dim strCurrentDate As String    'i.e. latest read data
Dim strCurrentCompetition As String 'i.e. latest read competition
Dim strHomeTeam As String
Dim strAwayTeam As String
Dim dtKickOff As Date

If wsFixtures.Cells(1, 1).value = vbNullString Then
    MsgBox prompt:="There are no fixtures in the workbook. This may be a result of trying to get the data while not connected to the internet. Please check your connection and try to get data from the web again."
    GetFixtureList = False
    Exit Function
End If

Dim i As Integer
For i = 1 To wsFixtures.UsedRange.Rows.Count
    If wsFixtures.Cells(i, 2).value = vbNullString Then  'it's either a date or a competition
        If wsFixtures.Cells(i + 1, 2).value = vbNullString Then
            strCurrentDate = wsFixtures.Cells(i, 1)
        Else
            strCurrentCompetition = wsFixtures.Cells(i, 1)
        End If
    Else ' row of fixture data
        strHomeTeam = wsFixtures.Cells(i, 1)
        strAwayTeam = wsFixtures.Cells(i, 3)
        dtKickOff = wsFixtures.Cells(i, 4)
        
        Dim NextMatch As cFixture
        Set NextMatch = New cFixture
        
       
        NextMatch.MatchDate = strCurrentDate
        NextMatch.Competition = strCurrentCompetition
        NextMatch.HomeTeam = strHomeTeam
        NextMatch.AwayTeam = strAwayTeam
        NextMatch.KickOff = dtKickOff
            
        ColFixtures.Add NextMatch

    End If
Next i
GetFixtureList = True
End Function

'==============================
'Name:GetWebData
'Purpose: Gets the data for fixtures and league tables from the web.
'         Subroutine manages the calling of all of the web queries.
'==============================

Private Sub GetWebData()
Dim wsFixtures As Worksheet
Dim wsPremierLeagueTable As Worksheet
Dim wsChampionshipTable As Worksheet
Dim wsLeagueOneTable As Worksheet
Dim wsLeagueTwoTable As Worksheet
Dim wsMenu As Worksheet
Dim booWebQuerySucceeded As Boolean

booWebQuerySucceeded = True
Set wsMenu = Application.ThisWorkbook.Sheets("Menu")
Set wsFixtures = Application.ThisWorkbook.Sheets("Fixtures")
Set wsPremierLeagueTable = Application.ThisWorkbook.Sheets("Premier_League_Table")
Set wsChampionshipTable = Application.ThisWorkbook.Sheets("Championship_Table")
Set wsLeagueOneTable = Application.ThisWorkbook.Sheets("League_One_Table")
Set wsLeagueTwoTable = Application.ThisWorkbook.Sheets("League_Two_Table")

Dim strFixtures As String
Dim strPremierLeagueTable As String
Dim strChampionshipTable As String
Dim strLeagueOneTable As String
Dim strLeagueTwoTable As String

strFixtures = "http://www.football.co.uk/fixtures"
strPremierLeagueTable = "http://www.football.co.uk/league-tables/premier-league"
strChampionshipTable = "http://www.football.co.uk/league-tables/championship"
strLeagueOneTable = "http://www.football.co.uk/league-tables/league-1"
strLeagueTwoTable = "http://www.football.co.uk/league-tables/league-2"

WebQuery wsFixtures, strFixtures, "qryFixtures", "1", booWebQuerySucceeded
    If booWebQuerySucceeded = False Then
    wsMenu.Select
     Exit Sub
    End If
    
WebQuery wsPremierLeagueTable, strPremierLeagueTable, "qryPremierLeague", "1", booWebQuerySucceeded

If booWebQuerySucceeded = False Then
    wsMenu.Select
     Exit Sub
    End If

WebQuery wsChampionshipTable, strChampionshipTable, "qryChampionship", "1", booWebQuerySucceeded

If booWebQuerySucceeded = False Then
    wsMenu.Select
     Exit Sub
    End If

WebQuery wsLeagueOneTable, strLeagueOneTable, "qryLeagueOne", "1", booWebQuerySucceeded

If booWebQuerySucceeded = False Then
    wsMenu.Select
     Exit Sub
    End If

WebQuery wsLeagueTwoTable, strLeagueTwoTable, "qryLeagueTwo", "1", booWebQuerySucceeded

If booWebQuerySucceeded = False Then
   
     Exit Sub
    End If

wsMenu.Select
End Sub

'==============================
'Name:WebQuery
'Purpose: Gets data from the web and stores it in a worksheet
'
'==============================

Private Sub WebQuery(ByVal wsDestination As Worksheet, _
        ByVal strURL As String, _
        ByVal qtName As String, _
        ByVal qtWebTables As String, _
        ByRef booWebQuerySucceeded As Boolean)

Dim qt As QueryTable
On Error GoTo errHandler
wsDestination.Visible = xlSheetVisible
wsDestination.Select
wsDestination.UsedRange.ClearContents
Set qt = wsDestination.QueryTables.Add(Connection:="URL;" & strURL, Destination:=wsDestination.Range("A1"))
qt.RefreshOnFileOpen = False
qt.Name = qtName
qt.FieldNames = True
qt.WebSelectionType = xlSpecifiedTables
qt.WebTables = qtWebTables
qt.Refresh BackgroundQuery:=False

wsDestination.Visible = xlSheetHidden
errHandler:
    If InStr(1, Err.Description, "Cannot locate the Internet server or proxy", vbTextCompare) > 0 Then
    MsgBox prompt:="Error getting data from the internet. Please check your connection. Macro stopped"
    booWebQuerySucceeded = False
    Exit Sub
    End If
    
    
End Sub

'==============================
'Name:ClearDownReport
'Purpose: Deletes the existing sheet and creates a new one. Preferred to deleting the data
'         because the used range will always be properly reset.
'==============================

Private Sub ClearDownReport()
Dim wsMenu As Worksheet
Set wsMenu = Application.ThisWorkbook.Worksheets("Menu")
Dim wsFixturesReport As Worksheet
Set wsFixturesReport = Application.ThisWorkbook.Sheets("Fixtures_Report")
Dim wsNewFixturesReport As Worksheet

Application.DisplayAlerts = False
wsFixturesReport.Delete
Application.DisplayAlerts = True
Set wsNewFixturesReport = ThisWorkbook.Worksheets.Add
wsNewFixturesReport.Move after:=ThisWorkbook.Worksheets("Menu")

wsNewFixturesReport.Name = "Fixtures_Report"
wsNewFixturesReport.Activate
ActiveWindow.DisplayGridlines = False
wsNewFixturesReport.Columns("A").ColumnWidth = 30
wsNewFixturesReport.Columns("B").ColumnWidth = 5
wsNewFixturesReport.Columns("C").ColumnWidth = 30
wsNewFixturesReport.Columns("D").ColumnWidth = 10
wsNewFixturesReport.Columns("D").NumberFormat = "hh:mm"

With wsNewFixturesReport.PageSetup
    .FitToPagesWide = 1
    .LeftMargin = Application.CentimetersToPoints(2)
    .RightMargin = Application.CentimetersToPoints(1)
End With

wsMenu.Select

End Sub

'==============================
'Name:CreateTeamFilter
'Purpose: Gets data from the filter worksheet and stores it in an array
'         For use in filtering fixtures. For example, to just London teams.
'==============================

Private Function CreateTeamFilter(ByRef arrTeamFilter() As String, ByVal wsSheet As Worksheet) As Boolean
Dim i As Long
On Error GoTo errHandler
ReDim Preserve arrTeamFilter(1 To wsSheet.UsedRange.Rows.Count)
For i = 1 To wsSheet.UsedRange.Rows.Count
    arrTeamFilter(i) = wsSheet.Cells(i, 1).value
Next i

CreateTeamFilter = True
Exit Function

errHandler:
CreateTeamFilter = False
End Function

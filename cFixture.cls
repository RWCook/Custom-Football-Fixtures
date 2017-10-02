VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMatchDate As String
Private pCompetition As String
Private pHomeTeam As String
Private pAwayTeam As String
Private pKickOff As Date

Public Property Get MatchDate() As String
    MatchDate = pMatchDate
End Property

Public Property Let MatchDate(ByVal value As String)
    pMatchDate = value
End Property

Public Property Get Competition() As String
    Competition = pCompetition
End Property

Public Property Let Competition(ByVal value As String)
    pCompetition = value
End Property

Public Property Get HomeTeam() As String
    HomeTeam = pHomeTeam
End Property

Public Property Let HomeTeam(ByVal value As String)
    pHomeTeam = value
End Property

Public Property Get AwayTeam() As String
    AwayTeam = pAwayTeam
End Property

Public Property Let AwayTeam(ByVal value As String)
    pAwayTeam = value
End Property

Public Property Get KickOff() As Date
    KickOff = pKickOff
End Property

Public Property Let KickOff(ByVal value As Date)
    pKickOff = value
End Property

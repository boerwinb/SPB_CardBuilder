Option Explicit On
Option Strict On

Imports System.Threading

Public Class frmMain

    ''' <summary>
    ''' data structures
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure MasterStat
        Dim birthYear As Integer
        Dim birthMonth As Integer
        Dim birthDay As Integer
        Dim birthCountry As String
        Dim birthState As String
        Dim birthCity As String
        Dim nameFirst As String
        Dim nameLast As String
        Dim Bats As String
        Dim throws As String
        Dim debut As String
        Dim college As String
    End Structure

    Private Structure FieldingStat
        Dim games As Integer
        Dim position As Integer
        Dim putouts As Integer
        Dim assists As Integer
        Dim errors As Integer
        Dim dps As Integer
        Dim pbs As Integer
    End Structure

    Private Structure BatCard
        Dim player As String
        Dim Field As String
        Dim OBR As String
        Dim SP As String
        Dim HitRun As String
        Dim CD As String
        Dim Sac As String
        Dim Inj As String
        Dim Hit1bf As String
        Dim Hit1b7 As String
        Dim hit1b8 As String
        Dim hit1b9 As String
        Dim hit2b7 As String
        Dim hit2b8 As String
        Dim hit2b9 As String
        Dim hit3b8 As String
        Dim hitHR As String
        Dim k As String
        Dim W As String
        Dim HPB As String
        Dim Out As String
        Dim Cht As String
        Dim BD As String
    End Structure

    Private Structure PitCard
        Dim player As String
        Dim Field As String
        Dim pbr As String
        Dim SR As String
        Dim RR As String
        Dim Hit1bf As String
        Dim Hit1b7 As String
        Dim hit1b8 As String
        Dim hit1b9 As String
        Dim bk As String
        Dim k As String
        Dim W As String
        Dim PB As String
        Dim wp As String
        Dim Out As String
        Dim StartRel As String
        Dim throwRating As String
        Dim BattingCard As String
    End Structure

    Private Sub cmdBuildCards_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBuildCards.Click
        Me.Cursor = Cursors.WaitCursor
        CreateBattingCards()
        CreatePitchingCards()
        Me.Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' builds the pitching cards
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreatePitchingCards()
        Dim yearStart As Integer
        Dim yearEnd As Integer
        Dim league As String
        Dim sqlQuery As String
        Dim dupPlayers As String

        Dim atBats As Integer
        Dim homeRuns As Integer
        Dim hits As Integer

        'Dim rsPitching As ADODB.Recordset
        Dim dsPitching As DataSet = Nothing
        'Dim rsBatting As ADODB.Recordset
        Dim dsBatting As DataSet = Nothing
        'Dim rsFielding As ADODB.Recordset
        Dim dsFielding As DataSet = Nothing
        'Dim rsDups As ADODB.Recordset
        Dim dsDups As DataSet = Nothing
        'Dim rsIP As ADODB.Recordset
        Dim dsIP As DataSet = Nothing
        'Dim cn As ADODB.Connection
        Dim Batter As New BatStat
        Dim Pitcher As New PitStat
        Dim Master As MasterStat
        Dim Card As PitCard
        Dim colTeamGames As Collection

        Dim leagueBattersFaced As Integer
        Dim cumulativeBFP As Integer

        Dim singles As Integer
        Dim strikeouts As Integer
        Dim walks As Integer
        Dim h1B7s As Integer
        Dim h1B8s As Integer
        Dim h1B9s As Integer
        Dim wildPitches As Integer
        Dim passedBalls As Integer
        Dim balks As Integer
        Dim outs As Integer
        Dim totalNumbers As Integer
        Dim birthState As String
        Dim DOB As String

        Dim wpFactor As Double

        Dim filPitOut As Integer

        Dim DataAccess As New clsDataAccess("lahman")
        Dim FACDataAccess As New clsDataAccess("fac")
        
        Try
            If txtYear.Text = Nothing Then
                yearStart = 1901
                yearEnd = Year(Now) - 1
            Else
                yearStart = CInt(txtYear.Text)
                yearEnd = yearStart
            End If
            For i As Integer = yearStart To yearEnd
                If Dir(GetAppLocation() & "\" & i.ToString, vbDirectory) = Nothing Then
                    My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\" & i.ToString)
                End If

                colTeamGames = New Collection
                DataAccess.GetTeamGames(colTeamGames, i.ToString)
                For leagueID As Integer = 1 To 2
                    
                    league = IIF(leagueID = 1, "AL", "NL")
                    'Find dups
                    sqlQuery = "SELECT playerid FROM pitching WHERE yearid = " & i.ToString & " AND " & _
                                  "lgid = '" & league & "' AND stint = 2"
                    'rsDups.Open(sqlQuery, cn)
                    dsDups = DataAccess.ExecuteDataSet(sqlQuery)
                    dupPlayers = ""
                    'While Not rsDups.EOF
                    '    dupPlayers += rsDups.Fields("playerid").Value.ToString & "|"
                    '    rsDups.MoveNext()
                    'End While
                    'rsDups.Close()
                    For Each dr As DataRow In dsDups.Tables(0).Rows
                        dupPlayers += dr.Item("playerid").ToString & "|"
                    Next dr

                    'Determine PB Rating thresholds for given league
                    sqlQuery = "SELECT sum(bfp) as totalbfp FROM pitching WHERE lgid = '" & _
                                league & "' AND " & _
                                "yearid = " & i.ToString
                    'rsIP.Open(sqlQuery, cn)
                    'leagueIpOuts = CInt(rsIP.Fields("totalip").Value)
                    'rsIP.Close()
                    dsIP = DataAccess.ExecuteDataSet(sqlQuery)
                    leagueBattersFaced = CInt(dsIP.Tables(0).Rows(0).Item("totalbfp"))

                    Call Pitcher.DetermineThresholds(leagueBattersFaced)

                    Call Pitcher.DetermineLeagueAvgs(league, i, DataAccess)

                    'Process each pitcher in league
                    cumulativeBFP = 0
                    'sqlQuery = String.Format("SELECT * from pitching a, master b, xb_allowed_{0}_{1} c where a.playerid = b.playerid and a.playerid = c.lahmanid and a.teamid = c.teamid and " & _
                    '                    "a.yearid = '{2}' and a.lgid = '{0}' and a.h > 0 order by CAST(c.""2b"" + c.""3b"" + c.hr as DOUBLE)/c.h asc", league, i.ToString.Substring(2), i.ToString)
                    sqlQuery = "SELECT * FROM pitching INNER JOIN people ON " & _
                              "pitching.playerid = people.playerid WHERE pitching.yearid = " & i.ToString & _
                              " AND " & "pitching.lgid = '" & league & "' and pitching.h > 0 ORDER BY pitching.era ASC"
                    '" AND " & "pitching.lgid = '" & league & "' and pitching.h > 0 ORDER BY (pitching.hr/pitching.h) ASC"
                    ' For Each dr As DataRow In dsStandings.Tables(0).Rows
                    'next dr
                    'With rsPitching
                    '.Open(sqlQuery, cn)
                    'While Not .EOF
                    dsPitching = DataAccess.ExecuteDataSet(sqlQuery)
                    For Each dr As DataRow In dsPitching.Tables(0).Rows
                        Pitcher.Clear()
                        Pitcher.team = dr.Item("teamid").ToString
                        Pitcher.wins = CInt(dr.Item("w"))
                        Pitcher.losses = CInt(dr.Item("l"))
                        Pitcher.saves = CInt(dr.Item("sv"))
                        Pitcher.games = CInt(dr.Item("g"))
                        Pitcher.gs = CInt(dr.Item("gs"))
                        Pitcher.cg = CInt(dr.Item("cg"))
                        Pitcher.outs = CInt(dr.Item("ipouts"))
                        Pitcher.hits = CInt(dr.Item("h"))
                        Pitcher.hrs = CInt(dr.Item("hr"))
                        Pitcher.bb = CInt(dr.Item("bb"))
                        Pitcher.ibb = CheckField(dr.Item("ibb"), 0)
                        Pitcher.k = CInt(dr.Item("so"))
                        Pitcher.wp = CheckField(dr.Item("wp"), 0)
                        Pitcher.bk = CheckField(dr.Item("bk"), 0)
                        Pitcher.era = CheckField(dr.Item("era"), 0)
                        If i > 1915 AndAlso CheckField(dr.Item("bfp"), 0) > 0 Then
                            'Batters Faced was tracked after 1915
                            Pitcher.bfp = CInt(dr.Item("bfp"))
                        Else
                            'Otherwise, try to determine from other stats
                            Pitcher.bfp = CInt(Pitcher.outs * (26 / 27) + Pitcher.bb + Pitcher.hits - Pitcher.ibb)
                        End If

                        Dim normalization As Double = 162 / CInt(colTeamGames(Pitcher.team))
                        If dr.Item("namelast").ToString = "Carrara" Then
                            Master.nameLast = "Carrara"
                        End If
                        If CInt(Pitcher.outs * normalization) >= 45 Or _
                                (CInt(Pitcher.gs * normalization) >= 3 Or CInt(Pitcher.games * normalization) >= 10) Then
                            'Valid pitcher
                            'Fill master stats
                            Master.Bats = CheckField(dr.Item("bats"), "").Trim
                            If Master.Bats = Nothing Then
                                Master.Bats = "R"
                            ElseIf Master.Bats = "B" Then
                                Master.Bats = "S"
                            End If
                            Master.throws = CheckField(dr.Item("throws"), "").Trim
                            If Master.throws = Nothing Then
                                Master.throws = "R"
                            End If
                            Master.birthCity = CheckField(dr.Item("birthcity"), "")
                            Master.birthCountry = CheckField(dr.Item("birthcountry"), "")
                            Master.birthDay = CheckField(dr.Item("birthday"), 0)
                            Master.birthMonth = CheckField(dr.Item("birthmonth"), 0)
                            Master.birthState = CheckField(dr.Item("birthstate"), "")
                            Master.birthYear = CheckField(dr.Item("birthyear"), 0)
                            'Master.college = CheckField(dr.Item("college"), "")
                            Master.debut = CheckField(dr.Item("debut"), "")
                            Master.nameFirst = dr.Item("namefirst").ToString
                            Master.nameLast = dr.Item("namelast").ToString

                            'Fill fielding stats
                            sqlQuery = "SELECT * FROM fielding WHERE playerid = '" & _
                                        dr.Item("playerid").ToString & "' AND " & _
                                        "teamid = '" & Pitcher.team & "' AND " & _
                                        "yearid = " & i & " AND POS = 'P'"
                            'rsFielding.Open(sqlQuery, cn)
                            dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                            Card.Field = ""
                            'While Not rsFielding.EOF
                            For Each drField As DataRow In dsFielding.Tables(0).Rows
                                Pitcher.po = CheckField(drField.Item("po"), 0)
                                Pitcher.err = CheckField(drField.Item("e"), 0)
                                Pitcher.assists = CheckField(drField.Item("a"), 0)
                                Pitcher.dps = CheckField(drField.Item("dp"), 0)
                                Pitcher.games = CheckField(drField.Item("g"), 0)

                                Card.Field = Card.Field & Pitcher.GetFieldingLine(CInt(colTeamGames(Pitcher.team)))
                                'rsFielding.MoveNext()
                                'End While
                            Next drField
                            'rsFielding.Close()

                            If league = "NL" Or i < 1973 Then
                                'Figure out pitcher batting
                                atBats = 0
                                hits = 0
                                homeRuns = 0
                                sqlQuery = "SELECT ab, h, hr FROM batting WHERE playerid = '" & _
                                            dr.Item("playerid").ToString & "' AND " & _
                                            "teamid = '" & Pitcher.team & "' AND " & _
                                            "yearid = " & i.ToString
                                'rsBatting.Open(sqlQuery, cn)
                                dsBatting = DataAccess.ExecuteDataSet(sqlQuery)
                                For Each drBatting As DataRow In dsBatting.Tables(0).Rows
                                    'While Not rsBatting.EOF
                                    atBats = CInt(drBatting.Item("ab"))
                                    hits = CInt(drBatting.Item("h"))
                                    homeRuns = CInt(drBatting.Item("hr"))
                                    'rsBatting.MoveNext()
                                    'End While
                                Next drBatting
                                'rsBatting.Close()
                                Card.BattingCard = Pitcher.GetBattingCardNum(atBats, hits, homeRuns)
                            Else
                                'AL, DH RULE era default
                                Card.BattingCard = "4"
                            End If

                            Card.player = Master.nameFirst & " " & Master.nameLast

                            'PB rating
                            Call Pitcher.GetPBRating(cumulativeBFP, Pitcher.bfp)
                            Card.pbr = Pitcher.pbr
                            'SR rating
                            If Pitcher.gs >= Pitcher.games - Pitcher.gs Then
                                'Starter
                                Card.SR = CInt(Pitcher.era * 1.75 + (Pitcher.hits + Pitcher.bb) / _
                                          (Pitcher.gs + 0.5 * (Pitcher.games - Pitcher.gs))).ToString
                                If (Pitcher.games - Pitcher.gs) > 0 Then
                                    Card.RR = CInt(Val(Card.SR) / 2).ToString
                                Else
                                    Card.RR = "0"
                                End If
                                If Val(Card.SR) < 12 Then
                                    Card.SR = "12"
                                End If
                            Else
                                'Reliever
                                Card.RR = CInt(Pitcher.era * 1.75 + (Pitcher.hits + Pitcher.bb) / _
                                          (Pitcher.games - Pitcher.gs + 2 * (Pitcher.gs))).ToString
                                If Val(Card.RR) < 3 Then
                                    Card.RR = "3"
                                End If
                                If (Pitcher.gs) > 0 Then
                                    Card.SR = IIF(Pitcher.gs > 1 And Val(Card.RR) < 12, "12", Card.RR)
                                Else
                                    Card.SR = "0"
                                End If
                            End If
                            totalNumbers = 64
                            For j As Integer = 1 To 2
                                'The first pass of this loop is all that is needed. The second
                                'pass is used if it is determined the numbers will exceed 
                                '88 on the card

                                'First, determine WP, BK, PB
                                passedBalls = IIF(j = 1, 1, 0)
                                wildPitches = Pitcher.DetermineWPNumbers(totalNumbers)
                                balks = Pitcher.DetermineBKNumbers(totalNumbers)
                                wpFactor = (64 - wildPitches - balks - passedBalls) / 64

                                'Singles, Strikeouts, Walks
                                Call Pitcher.DeterminePitcherNumbers(singles, strikeouts, walks, _
                                            wpFactor, totalNumbers)

                                totalNumbers = singles + balks + strikeouts + walks + passedBalls + _
                                                    wildPitches
                                If totalNumbers <= 64 Then
                                    'exit loop
                                    j = 2
                                End If
                            Next j

                            'determine 1bf's
                            Select Case singles
                                Case 7 To 16
                                    Card.Hit1bf = GetNumber(1)
                                    singles = singles - 1
                                Case Is > 16
                                    Card.Hit1bf = GetNumber(2)
                                    singles = singles - 2
                                Case Else
                                    Card.Hit1bf = ""
                            End Select
                            h1B7s = 0
                            h1B8s = 0
                            h1B9s = 0
                            Call AssignHitFields(singles, h1B7s, h1B8s, h1B9s, "S")
                            Card.Hit1b7 = GetNumber(h1B7s)
                            Card.hit1b8 = GetNumber(h1B8s)
                            Card.hit1b9 = GetNumber(h1B9s)

                            Card.bk = GetNumber(balks)

                            'determine k's
                            Card.k = GetNumber(strikeouts)

                            'determine walk's
                            Card.W = GetNumber(walks)

                            Card.PB = GetNumber(passedBalls)
                            Card.wp = GetNumber(wildPitches)

                            'Determine Outs
                            If gCurrentNum = 0 Then
                                outs = 11
                            ElseIf gCurrentNum Mod 10 = 8 Then
                                'end of a base 8 series
                                outs = gCurrentNum + 3
                            Else
                                outs = gCurrentNum + 1
                            End If
                            Select Case outs
                                Case Is > 88
                                    Card.Out = ""
                                Case 88
                                    Card.Out = "88"
                                Case Else
                                    Card.Out = outs & "-88"
                            End Select
                            'Determine S-R
                            Card.StartRel = Pitcher.gs.ToString & "-" & (Pitcher.games - Pitcher.gs).ToString
                            'Determine throw
                            Card.throwRating = IIF(Master.throws = "R", "Right", "Left")
                            gCurrentNum = 0
                            If txtYear.Text <> Nothing Then
                                lblStatus.Text = Master.nameFirst & " " & Master.nameLast & _
                                              Space(5) & Pitcher.team
                                System.Windows.Forms.Application.DoEvents()
                                Thread.Sleep(1)
                            End If

                            If Dir(GetAppLocation() & "\" & i.ToString & "\" & FACDataAccess.GetLongTeamTranslation(Pitcher.team), _
                                                    vbDirectory) = Nothing Then
                                My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\" & i.ToString & "\" & _
                                                FACDataAccess.GetLongTeamTranslation(Pitcher.team))
                            End If


                            filPitOut = FreeFile()
                            FileOpen(filPitOut, GetAppLocation() & "\" & i.ToString & _
                                                    "\" & FACDataAccess.GetLongTeamTranslation(Pitcher.team) & "\" & Pitcher.team & _
                                                    "pitch.csv", OpenMode.Append)
                            If FileLen(GetAppLocation() & "\" & i.ToString & "\" & _
                                        FACDataAccess.GetLongTeamTranslation(Pitcher.team) & "\" & Pitcher.team & _
                                        "pitch.csv") = 0 Then
                                PrintLine(filPitOut, "Pitcher,Field,PBR,SR,RR,1BF,1B7,1B8,1B9,BK,K,W,PB,WP,OUT,S-R,Throw," & _
                                                "Bat,City,State,Birthdate")
                            End If

                            With Master
                                birthState = IIF(.birthCountry = "USA", .birthState, .birthCountry)
                                DOB = .birthMonth & "/" & .birthDay & "/" & .birthYear
                            End With
                            With Card
                                WriteLine(filPitOut, Master.nameFirst & " " & Master.nameLast, .Field, _
                                          .pbr, .SR, .RR, .Hit1bf, .Hit1b7, .hit1b8, .hit1b9, .bk, .k, .W, .PB, _
                                          .wp, .Out, .StartRel, .throwRating, .BattingCard, FACDataAccess.GetLongTeamTranslation(Pitcher.team), _
                                          Master.birthCity, birthState, DOB)
                            End With
                            FileClose(filPitOut)
                        End If
                        cumulativeBFP += Pitcher.bfp
                    Next dr
                    '.MoveNext()
                    'End While
                    '.Close()
                    'End With
                    If txtYear.Text = Nothing Then
                        lblStatus.Text = "Year: " & i.ToString & " League: " & leagueID.ToString
                        System.Windows.Forms.Application.DoEvents()
                        Thread.Sleep(1)
                    End If
                Next leagueID
                Next i
            'If txtYear.Text <> Nothing Then
            '    Debug.Print(dupPlayers)
            'End If
        Catch ex As Exception
            Call MsgBox("CreatePitchingCards " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' builds the batting cards
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateBattingCards()
        Dim yearStart As Integer
        Dim yearEnd As Integer
        Dim league As String
        Dim sqlQuery As String
        Dim dupPlayers As String
        Dim filBatOut As Integer

        'Dim rsBatting As ADODB.Recordset
        Dim dsBatting As DataSet = Nothing
        'Dim rsFielding As ADODB.Recordset
        Dim dsFielding As DataSet = Nothing
        'Dim rsDups As ADODB.Recordset
        Dim dsDups As DataSet = Nothing
        'Dim rsTeams As ADODB.Recordset
        Dim dsTeams As DataSet = Nothing
        'Dim cn As ADODB.Connection
        Dim Batter As New BatStat
        Dim Master As MasterStat
        Dim Card As New BatCard
        Dim count As Integer
        Dim pGames As Integer
        Dim ofGames As Integer
        Dim infGames As Integer
        Dim dhGames As Integer
        Dim phGames As Integer
        Dim prGames As Integer
        Dim ofPosition As String
        Dim infPosition As String
        Dim dhPosition As String
        Dim colTeamGames As Collection
        Dim filTrades As Integer
        Dim singleAdjustment As Double
        Dim kAdjustment As Double
        Dim bbAdjustment As Double
        Dim colPitchersAL As New Collection
        Dim colPitchersNL As New Collection
        Dim pitcherIds As String
        Dim birthState As String
        Dim DOB As String

        Dim DataAccess As New clsDataAccess("lahman")
        Dim FACDataAccess As New clsDataAccess("fac")

        Try
            If txtYear.Text = Nothing Then
                yearStart = 1901
                yearEnd = Year(Now) - 1
            Else
                yearStart = CInt(txtYear.Text)
                yearEnd = yearStart
            End If
            For i As Integer = yearStart To yearEnd
                If Dir(GetAppLocation() & "\" & i.ToString, vbDirectory) <> Nothing Then
                    My.Computer.FileSystem.DeleteDirectory(GetAppLocation() & "\" & i.ToString, _
                            FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If
                My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\" & i.ToString)
                colTeamGames = New Collection
                DataAccess.GetTeamGames(colTeamGames, CStr(i))

                'keep track of players that were traded during the season
                filTrades = FreeFile()
                FileOpen(filTrades, GetAppLocation() & "\" & i.ToString & _
                                                    "trades.txt", OpenMode.Output)
                For leagueIndex As Integer = 1 To 2

                    league = IIF(leagueIndex = 1, "AL", "NL")
                    'Find dups
                    sqlQuery = "SELECT playerid FROM batting WHERE yearid = " & i.ToString & " AND " & _
                                  "lgid = '" & league & "' AND stint > 1"
                    'rsDups.Open(sqlQuery, cn)
                    dsDups = DataAccess.ExecuteDataSet(sqlQuery)
                    dupPlayers = ""
                    'While Not rsDups.EOF
                    For Each dr As DataRow In dsDups.Tables(0).Rows
                        dupPlayers += dr.Item("playerid").ToString & "|"
                    Next dr
                    'rsDups.MoveNext()
                    'End While
                    'rsDups.Close()

                    sqlQuery = "SELECT * FROM batting INNER JOIN people ON " & _
                              "batting.playerid = people.playerid WHERE batting.yearid = " & i.ToString & _
                              " AND " & "batting.lgid = '" & league & "' " & _
                              IIF(i >= 1974, "AND ab + bb + hbp > 0 ", "") & _
                              "ORDER BY batting.ab DESC"
                    Batter.DetermineLeagueAvgs(league, i, DataAccess)
                    singleAdjustment = Batter.adjSingle
                    kAdjustment = Batter.adjK
                    bbAdjustment = Batter.adjBB
                    'With rsBatting
                    '    .Open(sqlQuery, cn)
                    dsBatting = DataAccess.ExecuteDataSet(sqlQuery)
                    'While Not .EOF
                    For Each dr As DataRow In dsBatting.Tables(0).Rows
                        Batter.Clear()
                        Batter.curTeam = dr.Item("teamid").ToString
                        If InStr(dupPlayers, dr.Item("playerid").ToString) > 0 Then
                            'Have dups query all teams
                            If i >= 1974 Then
                                sqlQuery = "SELECT * FROM batting WHERE yearid = " & i.ToString & " AND " & _
                                                     "lgid = '" & league & "' AND  playerid = '" & _
                                                     dr.Item("playerid").ToString & "' " & _
                                                    "AND ab + bb + hbp > 0"
                            Else
                                sqlQuery = "SELECT * FROM batting WHERE yearid = " & i.ToString & " AND " & _
                                                    "lgid = '" & league & "' AND  playerid = '" & _
                                                    dr.Item("playerid").ToString & "'"
                            End If

                            'rsDups.Open(sqlQuery, cn)
                            dsDups = DataAccess.ExecuteDataSet(sqlQuery)

                            'While Not rsDups.EOF
                            For Each drDup As DataRow In dsDups.Tables(0).Rows
                                PrintLine(filTrades, dr.Item("namefirst").ToString & " " & _
                                            dr.Item("namelast").ToString & " " & drDup.Item("teamid").ToString & _
                                            " " & drDup.Item("g").ToString)
                                Batter.games += CInt(drDup.Item("g"))
                                Batter.ab += CheckField(drDup.Item("ab"), 0)
                                Batter.runs += CheckField(drDup.Item("r"), 0)
                                Batter.hits += CheckField(drDup.Item("h"), 0)
                                Batter.doubles += CheckField(drDup.Item("d"), 0)
                                Batter.triples += CheckField(drDup.Item("t"), 0)
                                Batter.hrs += CheckField(drDup.Item("hr"), 0)
                                Batter.sb += CheckField(drDup.Item("sb"), 0)
                                Batter.bb += CheckField(drDup.Item("bb"), 0)
                                Batter.k += CheckField(drDup.Item("so"), 0)
                                Batter.ibb += CheckField(drDup.Item("ibb"), 0)
                                Batter.hb += CheckField(drDup.Item("hbp"), 0)
                                Batter.sh += CheckField(drDup.Item("sh"), 0)
                                Batter.sf += CheckField(drDup.Item("sf"), 0)
                                count += 1
                            Next drDup
                            'rsDups.MoveNext()
                            'End While
                            'rsDups.Close()
                        Else
                            Batter.games = CInt(dr.Item("g"))
                            Batter.ab = CheckField(dr.Item("ab"), 0)
                            Batter.runs = CheckField(dr.Item("r"), 0)
                            Batter.hits = CheckField(dr.Item("h"), 0)
                            Batter.doubles = CheckField(dr.Item("d"), 0)
                            Batter.triples = CheckField(dr.Item("t"), 0)
                            Batter.hrs = CheckField(dr.Item("hr"), 0)
                            Batter.sb = CheckField(dr.Item("sb"), 0)
                            Batter.bb = CheckField(dr.Item("bb"), 0)
                            Batter.k = CheckField(dr.Item("so"), 0)
                            Batter.ibb = CheckField(dr.Item("ibb"), 0)
                            Batter.hb = CheckField(dr.Item("hbp"), 0)
                            Batter.sh = CheckField(dr.Item("sh"), 0)
                            Batter.sf = CheckField(dr.Item("sf"), 0)
                        End If


                        'Fill master stats
                        Master.Bats = CheckField(dr.Item("bats"), "").Trim
                        If Master.Bats = Nothing Then
                            Master.Bats = "R"
                        ElseIf Master.Bats = "B" Then
                            Master.Bats = "S"
                        End If
                        Master.throws = CheckField(dr.Item("throws"), "").Trim
                        If Master.throws = Nothing Then
                            Master.throws = "R"
                        End If
                        Master.birthCity = CheckField(dr.Item("birthcity"), "")
                        Master.birthCountry = CheckField(dr.Item("birthcountry"), "")
                        Master.birthDay = CheckField(dr.Item("birthday"), 0)
                        Master.birthMonth = CheckField(dr.Item("birthmonth"), 0)
                        Master.birthState = CheckField(dr.Item("birthstate"), "")
                        Master.birthYear = CheckField(dr.Item("birthyear"), 0)
                        'Master.college = CheckField(dr.Item("college"), "")
                        Master.debut = CheckField(dr.Item("debut"), "")
                        Master.nameFirst = dr.Item("namefirst").ToString
                        Master.nameLast = dr.Item("namelast").ToString

                        Card.Field = ""
                        Card.CD = "0"
                        ofGames = 0
                        infGames = 0
                        dhGames = 0
                        phGames = 0
                        prGames = 0
                        pGames = 0
                        infPosition = ""
                        dhPosition = ""
                        ofPosition = ""
                        'Determine games pitched. This value will help determine if the player is primarily a
                        'pitcher. For this version, we want to eliminate them from the batter card set.
                        sqlQuery = "SELECT * FROM fielding WHERE playerid = '" & _
                                     dr.Item("playerid").ToString & "' AND " & _
                                     "teamid = '" & Batter.curTeam & "' AND " & _
                                     "yearid = " & i.ToString & " AND " & _
                                     "pos = 'P'"
                        'rsFielding.Open(sqlQuery, cn)
                        dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                        'If Not rsFielding.EOF Then
                        If dsFielding.Tables(0).Rows.Count > 0 Then
                            Batter.curPos = "P"
                            pGames = CheckField(dsFielding.Tables(0).Rows(0).Item("g"), 0)
                        End If
                        'rsFielding.Close()

                        'Handle OF fielding stats
                        sqlQuery = "SELECT * FROM fielding WHERE playerid = '" & _
                                    dr.Item("playerid").ToString & "' AND " & _
                                    "teamid = '" & Batter.curTeam & "' AND " & _
                                    "yearid = " & i.ToString & " AND " & _
                                    IIF(i >= 2008, "pos in ('LF','CF','RF','OF')", "pos = 'OF'")
                        'rsFielding.Open(sqlQuery, cn)
                        dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                        'While Not rsFielding.EOF
                        For Each drField As DataRow In dsFielding.Tables(0).Rows
                            Batter.po += CheckField(drField.Item("po"), 0)
                            Batter.err += CheckField(drField.Item("e"), 0)
                            Batter.assists += CheckField(drField.Item("a"), 0)
                            Batter.dps += CheckField(drField.Item("dp"), 0)
                            Batter.posGames += CheckField(drField.Item("g"), 0)
                            Batter.curPos = "OF"
                        Next drField
                        '    rsFielding.MoveNext()
                        'End While
                        If i >= 1996 Then
                            ofGames = Math.Min(Batter.posGames, Batter.games)
                            Batter.posGames = ofGames
                        Else
                            ofGames = Batter.posGames
                        End If
                        'rsFielding.Close()
                        If Batter.curPos = "OF" Then
                            ofPosition = Batter.GetFieldingLine(Card.CD, CInt(colTeamGames(Batter.curTeam)))
                        End If

                        'Fill fielding stats
                        sqlQuery = "SELECT * FROM fielding WHERE playerid = '" & _
                                    dr.Item("playerid").ToString & "' AND " & _
                                    "teamid = '" & Batter.curTeam & "' AND " & _
                                    "yearid = " & i.ToString & " AND " & _
                                    "pos not in ('LF','CF','RF','OF','P') " & _
                                    "ORDER BY g DESC"
                        dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                        For Each drField As DataRow In dsFielding.Tables(0).Rows
                            Batter.po = CheckField(drField.Item("po"), 0)
                            Batter.err = CheckField(drField.Item("e"), 0)
                            Batter.assists = CheckField(drField.Item("a"), 0)
                            Batter.dps = CheckField(drField.Item("dp"), 0)
                            Batter.posGames = CheckField(drField.Item("g"), 0)
                            Batter.curPos = drField.Item("pos").ToString
                            infGames += Batter.posGames
                            infPosition += Batter.GetFieldingLine(Card.CD, CInt(colTeamGames(Batter.curTeam)))
                        Next drField

                        'Fill DH stats
                        sqlQuery = "SELECT * FROM appearances WHERE playerid = '" & _
                                    dr.Item("playerid").ToString & "' AND " & _
                                    "teamid = '" & Batter.curTeam & "' AND " & _
                                    "yearid = " & i.ToString & " AND " & _
                                    "g_dh >= 2 "
                        dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                        For Each drField As DataRow In dsFielding.Tables(0).Rows
                            Batter.po = 0
                            Batter.err = 0
                            Batter.assists = 0
                            Batter.dps = 0
                            Batter.posGames = CheckField(drField.Item("g_dh"), 0)
                            Batter.curPos = "DH"
                            dhGames += Batter.posGames
                            dhPosition += Batter.GetFieldingLine(Card.CD, CInt(colTeamGames(Batter.curTeam)))
                        Next drField

                        'Determine PH, PR stats
                        sqlQuery = "SELECT * FROM appearances WHERE playerid = '" & _
                                    dr.Item("playerid").ToString & "' AND " & _
                                    "teamid = '" & Batter.curTeam & "' AND " & _
                                    "yearid = " & i.ToString
                        dsFielding = DataAccess.ExecuteDataSet(sqlQuery)
                        For Each drField As DataRow In dsFielding.Tables(0).Rows
                            phGames += CheckField(drField.Item("g_ph"), 0)
                            prGames += CheckField(drField.Item("g_pr"), 0)
                        Next drField

                        Dim fieldLines() As String = {ofPosition, infPosition, dhPosition}
                        Dim fieldGames() As Integer = {ofGames, infGames, dhGames}
                        Array.Sort(fieldGames, fieldLines)
                        Array.Reverse(fieldLines)

                        For j As Integer = 0 To 2
                            'Print field lines in order most games to least games
                            Card.Field += fieldLines(j)
                        Next
                        'If ofGames > infGames Then
                        '    Card.Field = ofPosition & infPosition
                        'Else
                        '    Card.Field = infPosition & ofPosition
                        'End If

                        If pGames > ofGames + infGames + dhGames Then
                            If (ofGames + infGames + dhGames + phGames + prGames) < 20 Then
                                'Babe Ruth, Shohei, Brooks Kieschnick exception
                                Batter.curPos = "P"
                            End If
                            If league = "AL" Then
                                colPitchersAL.Add(dr.Item("playerid").ToString)
                            Else
                                colPitchersNL.Add(dr.Item("playerid").ToString)
                            End If
                        End If

                        Dim normalization As Double = 162 / CInt(colTeamGames(Batter.curTeam))
                        If (CInt(Batter.games * normalization) >= 20 And _
                                CInt((Batter.ab + Batter.bb + Batter.hb + Batter.sf + Batter.sh) * normalization) >= 20) Or _
                                CInt((Batter.ab + Batter.bb + Batter.hb + Batter.sf + Batter.sh) * normalization) >= 50 Then
                            'Valid batter
                            If Card.Field.Trim = Nothing And Batter.curPos <> "P" Then
                                Card.Field = "Pinch Hitter Only"
                            End If

                            'If Not (Card.Field.Trim = Nothing And Batter.curPos = "P") Then
                            If Batter.curPos <> "P" Then
                                FillBatterNumbers(Batter, Card, colTeamGames, singleAdjustment, kAdjustment, _
                                                    bbAdjustment, Master.Bats)
                                If txtYear.Text <> Nothing Then
                                    lblStatus.Text = Master.nameFirst & " " & Master.nameLast & Space(5) & Batter.curTeam
                                    System.Windows.Forms.Application.DoEvents()
                                    Thread.Sleep(1)
                                End If
                                If Dir(GetAppLocation() & "\" & i.ToString & "\" & FACDataAccess.GetLongTeamTranslation(Batter.curTeam), _
                                                    vbDirectory) = Nothing Then
                                    My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\" & i.ToString & "\" & _
                                                FACDataAccess.GetLongTeamTranslation(Batter.curTeam))
                                End If

                                filBatOut = FreeFile()
                                FileOpen(filBatOut, GetAppLocation() & "\" & i.ToString & _
                                              "\" & FACDataAccess.GetLongTeamTranslation(Batter.curTeam) & "\" & Batter.curTeam & _
                                              "bat.csv", OpenMode.Append)

                                If FileLen(GetAppLocation() & "\" & i.ToString & "\" & _
                                                    FACDataAccess.GetLongTeamTranslation(Batter.curTeam) & "\" & Batter.curTeam & _
                                                    "bat.csv") = 0 Then
                                    PrintLine(filBatOut, "Player,Field,OBR,SP,HITRUN,CD,SAC,INJ,1BF,1B7,1B8,1B9,2B7,2B8," & _
                                                    "2B9,3B8,HR,K,W,HPB,OUT,CHT,BD,City,State,Birthdate")
                                End If
                                With Master
                                    birthState = IIF(.birthCountry = "USA", .birthState, .birthCountry)
                                    DOB = .birthMonth & "/" & .birthDay & "/" & .birthYear
                                End With
                                With Card
                                    WriteLine(filBatOut, Master.nameFirst & " " & Master.nameLast, .Field, _
                                              .OBR, .SP, .HitRun, .CD, .Sac, .Inj, .Hit1bf, .Hit1b7, _
                                              .hit1b8, .hit1b9, .hit2b7, .hit2b8, .hit2b9, .hit3b8, .hitHR, _
                                              .k, .W, .HPB, .Out, .Cht, .BD, Master.birthCity, birthState, DOB)
                                End With
                                FileClose(filBatOut)
                            End If
                        End If
                        '    .MoveNext()
                        'End While
                        '    .Close()
                        'End With
                    Next dr
                    'Determine team pitching cards
                    If league = "NL" Or i < 1973 Or i >= 1997 Then
                        sqlQuery = "SELECT teamid,g from TEAMS WHERE yearid = " & i.ToString & " and lgid = '" & _
                                        league & "'"
                        'rsTeams.Open(sqlQuery, cn)
                        dsTeams = DataAccess.ExecuteDataSet(sqlQuery)
                        pitcherIds = BuildPitcherList(IIF(league = "NL", colPitchersNL, colPitchersAL))
                        'While Not rsTeams.EOF
                        For Each dr As DataRow In dsTeams.Tables(0).Rows
                            If i >= 1974 Then
                                sqlQuery = "SELECT sum(ab) as totalab, sum(r) as totalr, sum(h) as totalh, " & _
                                              "sum(d) as totald, sum(t) as totalt, sum(hr) as totalhr, sum(sb) as totalsb, " & _
                                              "sum(bb) as totalbb, sum(so) as totalso, sum(ibb) as totalibb, " & _
                                              "sum(hbp) as totalhbp, sum(sh) as totalsh, sum(sf) as totalsf " & _
                                              "FROM batting WHERE teamid = '" & dr.Item("teamid").ToString & "' " & _
                                              "AND yearid = " & i.ToString & " AND playerid in (" & pitcherIds & ") " & _
                                              "AND ab + bb + hbp > 0"
                            Else
                                sqlQuery = "SELECT sum(ab) as totalab, sum(r) as totalr, sum(h) as totalh, " & _
                                              "sum(d) as totald, sum(t) as totalt, sum(hr) as totalhr, sum(sb) as totalsb, " & _
                                              "sum(bb) as totalbb, sum(so) as totalso, sum(ibb) as totalibb, " & _
                                              "sum(hbp) as totalhbp, sum(sh) as totalsh, sum(sf) as totalsf " & _
                                              "FROM batting WHERE teamid = '" & dr.Item("teamid").ToString & "' " & _
                                              "AND yearid = " & i.ToString & " AND playerid in (" & pitcherIds & ")"
                            End If

                            'With rsBatting
                            '    .Open(sqlQuery, cn)
                            dsBatting = DataAccess.ExecuteDataSet(sqlQuery)
                            With dsBatting.Tables(0).Rows(0)
                                Batter.Clear()
                                Batter.curTeam = dr.Item("teamid").ToString
                                Batter.games = IIF(league = "AL" And i >= 1997, 9, CInt(dr.Item("g")))

                                Batter.ab = CInt(.Item("totalab"))
                                Batter.runs = CInt(.Item("totalr"))
                                Batter.hits = CInt(.Item("totalh"))
                                Batter.doubles = CInt(.Item("totald"))
                                Batter.triples = CInt(.Item("totalt"))
                                Batter.hrs = CInt(.Item("totalhr"))
                                Batter.sb = CInt(.Item("totalsb"))
                                Batter.bb = CInt(.Item("totalbb"))
                                Batter.k = CheckField(.Item("totalso"), 0)
                                Batter.ibb = CheckField(.Item("totalibb"), 0)
                                Batter.hb = CheckField(.Item("totalhbp"), 0)
                                Batter.sh = CInt(.Item("totalsh"))
                                Batter.sf = CheckField(.Item("totalsf"), 0)
                                '.Close()
                            End With
                            Card.Field = ""
                            Card.CD = "0"
                            Batter.curPos = "P"
                            FillBatterNumbers(Batter, Card, colTeamGames, singleAdjustment, kAdjustment, _
                                                     bbAdjustment, "S")
                            Card.Cht = "P"
                            If txtYear.Text <> Nothing Then
                                lblStatus.Text = "Team Batting Card for Pitchers" & Space(5) & Batter.curTeam
                                System.Windows.Forms.Application.DoEvents()
                                Thread.Sleep(1)
                            End If

                            If Dir(GetAppLocation() & "\" & i.ToString, vbDirectory) = Nothing Then
                                My.Computer.FileSystem.CreateDirectory(GetAppLocation() & "\" & i.ToString)
                            End If

                            filBatOut = FreeFile()
                            FileOpen(filBatOut, GetAppLocation() & "\" & i.ToString & "\" & "teamPitchBat.csv", _
                                                OpenMode.Append)

                            If FileLen(GetAppLocation() & "\" & i.ToString & "\" & "teamPitchBat.csv") = 0 Then
                                PrintLine(filBatOut, "Player,Field,OBR,SP,HITRUN,CD,SAC,INJ,1BF,1B7,1B8,1B9,2B7,2B8," & _
                                                "2B9,3B8,HR,K,W,HPB,OUT,CHT,BD,City,State,Birthdate")
                            End If

                            With Card
                                WriteLine(filBatOut, Batter.curTeam, .Field, _
                                          .OBR, .SP, .HitRun, .CD, .Sac, .Inj, .Hit1bf, .Hit1b7, _
                                          .hit1b8, .hit1b9, .hit2b7, .hit2b8, .hit2b9, .hit3b8, .hitHR, _
                                          .k, .W, .HPB, .Out, .Cht, .BD, "", "", "")
                            End With
                            FileClose(filBatOut)
                        Next dr
                        'rsTeams.MoveNext()
                        'End While
                        'rsTeams.Close()
                    End If
                    If txtYear.Text = Nothing Then
                        lblStatus.Text = "Year: " & i.ToString & Space(5) & "League: " & leagueIndex.ToString
                        System.Windows.Forms.Application.DoEvents()
                        Thread.Sleep(1)
                    End If
                Next leagueIndex
                    FileClose(filTrades)
                Next i

        Catch ex As Exception
            Call MsgBox("CreateBattingCards " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub FillBatterNumbers(ByRef Batter As BatStat, ByRef Card As BatCard, ByVal colTeamGames As Collection, _
                                ByRef singleAdjustment As Double, ByRef kAdjustment As Double, ByRef bbAdjustment As Double, _
                                ByVal batSide As String)
        Dim runAvg As String
        Dim totalNumbers As Integer
        Dim singleNumbers As Integer
        Dim doubleNumbers As Integer
        Dim tripleNumbers As Integer
        Dim homerunNumbers As Integer
        Dim strikeoutNumbers As Integer
        Dim walkNumbers As Integer
        Dim hpbNumbers As Integer
        Dim outStart As Integer
        Dim miscFactor As Double
        Dim battingFactor As Double
        Dim hitB7 As Integer
        Dim hitB8 As Integer
        Dim hitB9 As Integer
        Dim slugPct As Double

        'SB - extrapolate over 162 games
        Batter.sb = CInt(Batter.sb * 162 / Batter.games)
        Select Case Batter.sb
            Case 0
                Card.SP = "E"
            Case 1 To 9
                Card.SP = "D"
            Case 10 To 19
                Card.SP = "C"
            Case 20 To 29
                Card.SP = "B"
            Case 30 To 59
                Card.SP = "A"
            Case 60 To 89
                If Batter.games > 99 Then
                    Card.SP = "Y"
                Else
                    Card.SP = "A"
                End If
            Case Else
                If Batter.games > 99 Then
                    Card.SP = "Z"
                Else
                    Card.SP = "A"
                End If
        End Select

        'OBR - Calculate number of runs per times on base (excluding home runs)
        'However, base it mostly on SP
        With Batter
            If .hits + .bb + .hb - .hrs > 0 Then 'Avoid overflow
                runAvg = Format((.runs - .hrs) / (.hits + .bb + .hb - .hrs), "#.###")
            Else
                runAvg = ".3" 'Default to OBR C
            End If
        End With
        Select Case Card.SP
            Case "A", "Y", "Z", "B"
                Card.OBR = "A"
            Case "C"
                If Val(runAvg) > 0.4 Then
                    Card.OBR = "A"
                Else
                    Card.OBR = "B"
                End If
            Case "D"
                If Val(runAvg) > 0.375 Then
                    Card.OBR = "B"
                Else
                    Card.OBR = "C"
                End If
            Case "E"
                Select Case Val(runAvg)
                    Case 0 To 0.2
                        Card.OBR = "E"
                    Case 0.201 To 0.26
                        Card.OBR = "D"
                    Case 0.261 To 0.375
                        Card.OBR = "C"
                    Case Else
                        Card.OBR = "B"
                End Select
            Case Else
                Card.OBR = "C"
        End Select

        'Sacrifice hits, extrapolate over 162
        Batter.sh = CInt(Batter.sh * 162 / Batter.games)
        Select Case Batter.sh
            Case 0 To 1
                Card.Sac = "DD"
            Case 2 To 3
                Card.Sac = "CC"
            Case 4 To 6
                Card.Sac = "BB"
            Case Else
                Card.Sac = "AA"
        End Select

        'Inj rate
        Select Case CInt((Batter.games / CInt(colTeamGames(Batter.curTeam))) * 162)
            Case Is > 161
                Card.Inj = "0"
            Case 160 To 161
                Card.Inj = "1"
            Case 150 To 159
                Card.Inj = "2"
            Case 140 To 149
                Card.Inj = "3"
            Case 121 To 139
                Card.Inj = "4"
            Case 102 To 120
                Card.Inj = "5"
            Case 83 To 101
                Card.Inj = "6"
            Case 64 To 82
                Card.Inj = "7"
            Case 0 To 63
                Card.Inj = "8"
            Case Else
                Card.Inj = "8"
        End Select
        totalNumbers = 64
        For j As Integer = 1 To 2
            miscFactor = 64 / totalNumbers
            'Singles
            With Batter
                battingFactor = (.ab + .bb - .ibb + .hb + .sf) / 128
                singleNumbers = CInt(((.hits - .doubles - .triples - .hrs) / battingFactor - singleAdjustment) * miscFactor)
                If singleNumbers < 0 Then singleNumbers = 0
            End With



            'Extra base hits
            doubleNumbers = CInt((Batter.doubles / battingFactor) * miscFactor)
            tripleNumbers = CInt((Batter.triples / battingFactor) * miscFactor)
            homerunNumbers = CInt((Batter.hrs / battingFactor) * miscFactor)

            'Determine Strikeouts, Walks, Hit and Run, Hit Batters
            strikeoutNumbers = CInt((Batter.k / battingFactor - kAdjustment) * miscFactor)
            If strikeoutNumbers < 0 Then strikeoutNumbers = 0

            walkNumbers = CInt(((Batter.bb - Batter.ibb) / battingFactor - bbAdjustment) * miscFactor)
            If walkNumbers < 0 Then walkNumbers = 0

            hpbNumbers = CInt((Batter.hb / battingFactor) * miscFactor)


            totalNumbers = singleNumbers + doubleNumbers + tripleNumbers + homerunNumbers + strikeoutNumbers + _
                                    walkNumbers + hpbNumbers
            If totalNumbers <= 64 Then
                'exit loop
                j = 2
            End If
        Next j

        Card.Cht = batSide & IIF(homerunNumbers >= 4, "P", "N") 'P = Power, N = Normal

        'determine 1bf's
        Select Case singleNumbers
            Case 7 To 16
                Card.Hit1bf = GetNumber(1)
                singleNumbers -= 1
            Case Is > 16
                Card.Hit1bf = GetNumber(2)
                singleNumbers -= 2
            Case Else
                Card.Hit1bf = ""
        End Select
        hitB7 = 0
        hitB8 = 0
        hitB9 = 0
        Call AssignHitFields(singleNumbers, hitB7, hitB8, hitB9, Strings.Left(Card.Cht, 1))
        Card.Hit1b7 = GetNumber(hitB7)
        Card.hit1b8 = GetNumber(hitB8)
        Card.hit1b9 = GetNumber(hitB9)

        hitB7 = 0
        hitB8 = 0
        hitB9 = 0
        Call AssignHitFields(doubleNumbers, hitB7, hitB8, hitB9, Strings.Left(Card.Cht, 1))
        Card.hit2b7 = GetNumber(hitB7)
        Card.hit2b8 = GetNumber(hitB8)
        Card.hit2b9 = GetNumber(hitB9)

        Card.hit3b8 = GetNumber(tripleNumbers)
        Card.hitHR = GetNumber(homerunNumbers)

        'BD - determined initially from slugging percentage
        With Batter
            slugPct = CDbl(Format((.hits + .doubles + 2 * .triples + 3 * .hrs) / .ab, "0.###"))
        End With
        Select Case slugPct
            Case Is > 0.5
                Card.BD = "2"
            Case 0.4 To 0.499
                Card.BD = "1"
            Case Else
                Card.BD = "0"
        End Select
        Card.k = GetNumber(strikeoutNumbers)
        Select Case strikeoutNumbers
            Case 0 To 1
                Card.HitRun = "2"
            Case 2 To 3
                Card.HitRun = "1"
            Case Else
                Card.HitRun = "0"
        End Select
        Card.W = GetNumber(walkNumbers)
        Card.HPB = GetNumber(hpbNumbers)

        'Determine Outs
        If gCurrentNum = 0 Then
            outStart = 11
        ElseIf gCurrentNum Mod 10 = 8 Then
            'end of a base 8 series
            outStart = gCurrentNum + 3
        Else
            outStart = gCurrentNum + 1
        End If
        Select Case outStart
            Case Is > 88
                Card.Out = ""
            Case 88
                Card.Out = "88"
            Case Else
                Card.Out = outStart & "-88"
        End Select
        gCurrentNum = 0
    End Sub

    Public Structure foo
        Dim action As String
        Dim fname As String
        Dim fvalue As String
    End Structure
    
End Class



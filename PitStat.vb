Option Explicit On
Option Strict On

Public Class PitStat
    Private _Player As String
    Private _Team As String
    Private _Wins As Integer
    Private _Losses As Integer
    Private _Saves As Integer
    Private _Games As Integer
    Private _GS As Integer
    Private _CG As Integer
    Private _Outs As Integer
    Private _Hits As Integer
    Private _HR As Integer
    Private _BB As Integer
    Private _IBB As Integer
    Private _K As Integer
    Private _ERA As Double
    Private _BK As Integer
    Private _WP As Integer
    Private _HBP As Integer
    Private _PBR As String
    Private _BFP As Integer 'Batter's faced

    Private _PO As Integer
    Private _Err As Integer
    Private _Assists As Integer
    Private _DPS As Integer

    Private _28 As Double
    Private _27 As Double
    Private _26 As Double
    Private _25 As Double

    Private _SinglePct As Double
    Private _adjSingle As Double
    Private _adjK As Double
    Private _adjBB As Double
    Private _League As String
    Private _shPct As Double

    'Private Const AL_K_ADJUSTMENT As Double = 7.26 '8.26
    'Private Const AL_W_ADJUSTMENT As Double = 2.12 '3.62
    'Private Const NL_K_ADJUSTMENT As Double = 6.16 '7.66
    'Private Const NL_W_ADJUSTMENT As Double = 2.825 '3.825
    'Private Const SINGLE_ADJUSTMENT As Double = 11

    Public Sub Clear()
        _Player = ""
        _Team = ""
        _Wins = 0
        _Losses = 0
        _Saves = 0
        _Games = 0
        _GS = 0
        _CG = 0
        _Outs = 0
        _Hits = 0
        _HR = 0
        _BB = 0
        _IBB = 0
        _K = 0
        _BK = 0
        _WP = 0
        _HBP = 0
        _ERA = 0
        _PBR = ""
        _BFP = 0
        _PO = 0
        _Err = 0
        _Assists = 0
        _DPS = 0
    End Sub

    Public Property team() As String
        Get
            team = _Team
        End Get
        Set(ByVal value As String)
            _Team = value
        End Set
    End Property

    Public Property wins() As Integer
        Get
            wins = _Wins
        End Get
        Set(ByVal value As Integer)
            _Wins = value
        End Set
    End Property

    Public Property losses() As Integer
        Get
            losses = _Losses
        End Get
        Set(ByVal value As Integer)
            _Losses = value
        End Set
    End Property

    Public Property saves() As Integer
        Get
            saves = _Saves
        End Get
        Set(ByVal value As Integer)
            _Saves = value
        End Set
    End Property

    Public Property games() As Integer
        Get
            games = _Games
        End Get
        Set(ByVal value As Integer)
            _Games = value
        End Set
    End Property

    Public Property gs() As Integer
        Get
            gs = _GS
        End Get
        Set(ByVal value As Integer)
            _GS = value
        End Set
    End Property

    Public Property cg() As Integer
        Get
            cg = _CG
        End Get
        Set(ByVal value As Integer)
            _CG = value
        End Set
    End Property

    Public Property outs() As Integer
        Get
            outs = _Outs
        End Get
        Set(ByVal value As Integer)
            _Outs = value
        End Set
    End Property

    Public Property hits() As Integer
        Get
            hits = _Hits
        End Get
        Set(ByVal value As Integer)
            _Hits = value
        End Set
    End Property

    Public Property hrs() As Integer
        Get
            hrs = _HR
        End Get
        Set(ByVal value As Integer)
            _HR = value
        End Set
    End Property

    Public Property bb() As Integer
        Get
            bb = _BB
        End Get
        Set(ByVal value As Integer)
            _BB = value
        End Set
    End Property

    Public Property ibb() As Integer
        Get
            ibb = _IBB
        End Get
        Set(ByVal value As Integer)
            _IBB = value
        End Set
    End Property

    Public Property k() As Integer
        Get
            k = _K
        End Get
        Set(ByVal value As Integer)
            _K = value
        End Set
    End Property

    Public Property bk() As Integer
        Get
            bk = _BK
        End Get
        Set(ByVal value As Integer)
            _BK = value
        End Set
    End Property

    Public Property wp() As Integer
        Get
            wp = _WP
        End Get
        Set(ByVal value As Integer)
            _WP = value
        End Set
    End Property

    Public Property hbp() As Integer
        Get
            hbp = _HBP
        End Get
        Set(value As Integer)
            _HBP = value
        End Set
    End Property

    Public Property era() As Double
        Get
            era = _ERA
        End Get
        Set(ByVal value As Double)
            _ERA = value
        End Set
    End Property

    Public Property pbr() As String
        Get
            pbr = _PBR
        End Get
        Set(ByVal value As String)
            _PBR = value
        End Set
    End Property

    Public Property bfp() As Integer
        Get
            bfp = _BFP
        End Get
        Set(ByVal value As Integer)
            _BFP = value
        End Set
    End Property

    Public Property po() As Integer
        Get
            po = _PO
        End Get
        Set(ByVal value As Integer)
            _PO = value
        End Set
    End Property

    Public Property err() As Integer
        Get
            err = _Err
        End Get
        Set(ByVal value As Integer)
            _Err = value
        End Set
    End Property

    Public Property assists() As Integer
        Get
            assists = _Assists
        End Get
        Set(ByVal value As Integer)
            _Assists = value
        End Set
    End Property

    Public Property dps() As Integer
        Get
            dps = _DPS
        End Get
        Set(ByVal value As Integer)
            _DPS = value
        End Set
    End Property

    Public Property player() As String
        Get
            player = _Player
        End Get
        Set(ByVal Value As String)
            _Player = Value
        End Set
    End Property

    ''' <summary>
    ''' determines field ratings of pitcher
    ''' </summary>
    ''' <param name="games"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFieldingLine(ByVal games As Integer) As String
        Dim fieldLine As String = ""
        Dim percentValue As Double
        Dim errorRating As Integer
        Dim cdRating As String

        Try
            'Calculate percentage BEFORE PROJECTING OVER 162!
            If (_PO + _Assists + _Err) = 0 Then
                percentValue = 0
            Else
                percentValue = Val(Format((_PO + _Assists) / (_PO + _Assists + _Err), "0.###"))
            End If
            If _Games > 19 * games / 162 Then
                'Extrapolate assists, DPs over 162
                _Assists = CInt(Val(_Assists) * 162 / Val(_Games))
                _DPS = CInt(Val(_DPS) * 162 / Val(_Games))
            End If


            Select Case percentValue
                Case 1
                    errorRating = 0
                Case 0.985 To 0.999
                    errorRating = 1
                Case 0.975 To 0.984
                    errorRating = 2
                Case 0.965 To 0.974
                    errorRating = 3
                Case 0.955 To 0.964
                    errorRating = 4
                Case 0.945 To 0.954
                    errorRating = 5
                Case 0.935 To 0.944
                    errorRating = 6
                Case 0.925 To 0.934
                    errorRating = 7
                Case 0.915 To 0.924
                    errorRating = 8
                Case 0.905 To 0.914
                    errorRating = 9
                Case Else
                    errorRating = 10
            End Select

            If _Games > 19 * games / 162 And errorRating < 5 Then 'Only good fielders
                Select Case _DPS
                    Case 6 To 11
                        cdRating = "1"
                    Case Is > 11
                        cdRating = "2"
                    Case Else
                        cdRating = "0"
                End Select
            Else
                cdRating = "0"
            End If
            fieldLine = "E" & errorRating.ToString & " CD" & cdRating
        Catch ex As Exception
            Call MsgBox("GetFieldingLine " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return fieldLine
    End Function

    ''' <summary>
    ''' determines the hitting card for the pitcher
    ''' </summary>
    ''' <param name="atBats"></param>
    ''' <param name="hits"></param>
    ''' <param name="homeRuns"></param>
    ''' <returns></returns>
    ''' <remarks>these are not based on actual stats. One of 10 cards are chosen that
    ''' most closely resemble the actual stats</remarks>
    Public Function GetBattingCardNum(ByVal atBats As Integer, ByVal hits As Integer, _
                ByVal homeRuns As Integer) As String
        Dim battingAvg As Double
        Dim cardNumber As String = ""

        Try
            If atBats = 0 Then
                cardNumber = "1"
            Else
                'Calculate batting card number
                battingAvg = CDbl(Format(hits / atBats, "#.000"))

                If homeRuns > 0 Then
                    'use cards 8 thru 10
                    Select Case battingAvg
                        Case 0 To 0.151
                            cardNumber = "8"
                        Case 0.152 To 0.216
                            cardNumber = "9"
                        Case Else
                            cardNumber = "10"
                    End Select
                Else
                    'use 1 thru 7
                    Select Case battingAvg
                        Case 0 To 0.081
                            cardNumber = "1"
                        Case 0.082 To 0.122
                            cardNumber = "2"
                        Case 0.123 To 0.149
                            cardNumber = "3"
                        Case 0.15 To 0.178
                            cardNumber = "4"
                        Case 0.179 To 0.207
                            cardNumber = "5"
                        Case 0.208 To 0.239
                            cardNumber = "6"
                        Case Else
                            cardNumber = "7"
                    End Select
                End If
            End If
        Catch ex As Exception
            Call MsgBox("GetBattingCardNum " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return cardNumber
    End Function

    Public Sub DetermineThresholds(ByVal battersFaced As Integer)
        _28 = CDbl(Format(battersFaced * 0.05, "#####.##"))
        _27 = CDbl(Format(battersFaced * 0.15, "#####.##"))
        _26 = CDbl(Format(battersFaced * 0.45, "#####.##"))
        _25 = CDbl(Format(battersFaced * 0.85, "#####.##"))
    End Sub

    ''' <summary>
    ''' 
    '''determines league averages and K, W and Single factors. The pitching cards are built based on these factors
    ''' </summary>
    ''' <param name="league"></param>
    ''' <param name="year"></param>
    ''' <remarks></remarks>
    Public Sub DetermineLeagueAvgs(ByVal league As String, ByVal year As Integer, ByRef DataAccess As clsDataAccess)
        Dim outcomeRatio As Double
        Dim totalSingles As Integer
        'Dim kAdjustment As Double
        'Dim bbAdjustment As Double
        Dim totalH As Integer
        Dim totalD As Integer
        Dim totalT As Integer
        Dim totalHR As Integer
        Dim totalK As Integer
        Dim totalBB As Integer
        Dim totalIBBS As Integer
        Dim totalBFP As Integer
        Dim totalSH As Integer
        Dim sqlQuery As String
        'Dim rs As ADODB.Recordset
        Dim ds As DataSet = Nothing
        'Dim cn As ADODB.Connection


        Try
            'rs = New ADODB.Recordset
            'cn = New ADODB.Connection
            'cn.Open(GetLahmanConnectString)
            If year > 1915 Then
                'Batters Faced were tracked after 1915
                sqlQuery = "SELECT sum(bfp) as totalbfp FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                'rs.Open(sqlQuery, cn)
                ds = DataAccess.ExecuteDataSet(sqlQuery)
                totalBFP = CInt(ds.Tables(0).Rows(0).Item("totalbfp"))
            Else
                'Otherwise, try to determine from other stats
                sqlQuery = "SELECT sum(ipouts) as totalouts, sum(bb) as totalbb, sum(ibb) as totalibb, " & _
                                "sum(h) as totalhits FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                'rs.Open(sqlQuery, cn)
                ds = DataAccess.ExecuteDataSet(sqlQuery)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(CInt(.Item("totalouts")) * (26 / 27)) + CInt(.Item("totalbb")) + _
                                    CInt(.Item("totalhits")) '- CheckField(.item("totalibb"), 0)
                End With
            End If

            'rs.Close()

            sqlQuery = "SELECT sum(h) as totalhits, sum(d) as totalds, sum(t) as totalts, " & _
                                  "sum(hr) as totalhrs, sum(so) as totalks, sum(bb) as totalbbs, " & _
                                  "sum(ibb) as totalibb, sum(sh) as totalsh " & _
                                  "FROM batting WHERE lgid = '" & league & "' AND " & _
                                  "yearid = " & year.ToString

            'rs.Open(sqlQuery, cn)
            ds = DataAccess.ExecuteDataSet(sqlQuery)
            With ds.Tables(0).Rows(0)
                totalH = CInt(.Item("totalhits"))
                totalD = CInt(.Item("totalds"))
                totalT = CInt(.Item("totalts"))
                totalHR = CInt(.Item("totalhrs"))
                totalK = CheckField(.Item("totalks"), 0)
                totalBB = CInt(.Item("totalbbs"))
                totalIBBS = CheckField(.Item("totalibb"), 0)
                totalSH = CheckField(.Item("totalsh"), 0)
                'rs.Close()
            End With

            _League = league
            'kAdjustment = IIF(_League = "AL", AL_K_ADJUSTMENT, NL_K_ADJUSTMENT)
            'bbAdjustment = IIF(_League = "AL", AL_W_ADJUSTMENT, NL_W_ADJUSTMENT)

            totalSingles = totalH - totalD - totalT - totalHR
            _SinglePct = totalSingles / totalH
            outcomeRatio = (totalBFP - totalIBBS - totalSH) / 128
            _adjSingle = (totalSingles / outcomeRatio) / 2
            _adjK = (totalK / outcomeRatio) / 2
            _adjBB = ((totalBB - totalIBBS) / outcomeRatio) / 2
            _shPct = totalSH / totalBFP
        Catch ex As Exception
            Call MsgBox("DetermineLeagueAvgs " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' Determines K, W, Hit factors on pitcher card
    ''' </summary>
    ''' <param name="singleCardNumbers"></param>
    ''' <param name="kCardNumbers"></param>
    ''' <param name="bbCardNumbers"></param>
    ''' <param name="wpFactor"></param>
    ''' <param name="totalNumbers">The total numbers on the pitcher card. Used to handle
    ''' situations where the card numbers will exceed 88</param>
    ''' <remarks></remarks>
    Public Sub DeterminePitcherNumbers(ByRef singleCardNumbers As Integer, ByRef kCardNumbers As Integer, _
                                ByRef bbCardNumbers As Integer, ByVal wpFactor As Double, ByVal totalNumbers As Integer)
        Dim outcomeRatio As Double
        Dim singles As Double
        Dim strikeouts As Double
        Dim walks As Double
        Dim miscFactor As Double = 64 / totalNumbers 'Almost always 1.
        Dim numSH As Double
        
        Try
            numSH = _BFP * _shPct
            outcomeRatio = (_BFP - _IBB - _HBP - numSH) / 128
            singles = (_Hits * _SinglePct) / outcomeRatio - _adjSingle
            strikeouts = _K / outcomeRatio - _adjK
            walks = (_BB - _IBB) / outcomeRatio - _adjBB


            'Adjust numbers for PB ratings
            Select Case _PBR
                Case "2-9"
                    'singleCardNumbers = CInt((0.334 * singles + 6.71) * wpFactor * miscFactor)
                    'singleCardNumbers = CInt((0.334 * singles + 9.08) * wpFactor * miscFactor)
                    singleCardNumbers = CInt((0.334 * singles + 8.81) * wpFactor * miscFactor)
                    'kCardNumbers = CInt((0.334 * strikeouts + 5.63) * wpFactor * miscFactor)
                    kCardNumbers = CInt((0.334 * strikeouts + 6.62) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt((0.334 * walks + 4.5) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt((0.334 * walks + 3.83) * wpFactor * miscFactor)
                    bbCardNumbers = CInt((0.334 * walks + 3.03) * wpFactor * miscFactor)
                Case "2-8"
                    'singleCardNumbers = CInt((0.556 * singles + 4.48) * wpFactor * miscFactor)
                    'singleCardNumbers = CInt((0.556 * singles + 6.05) * wpFactor * miscFactor)
                    singleCardNumbers = CInt((0.556 * singles + 6.07) * wpFactor * miscFactor)
                    'kCardNumbers = CInt((0.556 * strikeouts + 3.76) * wpFactor * miscFactor)
                    kCardNumbers = CInt((0.556 * strikeouts + 4.48) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt((0.556 * walks + 3) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt((0.556 * walks + 2.4) * wpFactor * miscFactor)
                    bbCardNumbers = CInt((0.556 * walks + 2.09) * wpFactor * miscFactor)
                Case "2-7"
                    'singleCardNumbers = CInt((0.833 * singles + 1.68) * wpFactor * miscFactor)
                    singleCardNumbers = CInt((0.833 * singles + 1.88) * wpFactor * miscFactor)
                    'kCardNumbers = CInt((0.833 * strikeouts + 1.41) * wpFactor * miscFactor)
                    kCardNumbers = CInt((0.833 * strikeouts + 0.99) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt((0.833 * walks + 1.13) * wpFactor * miscFactor)
                    bbCardNumbers = CInt((0.833 * walks + 0.93) * wpFactor * miscFactor)
                Case "2-6"
                    'singleCardNumbers = CInt(((singles - 1.68) / 0.833) * wpFactor * miscFactor)
                    'singleCardNumbers = CInt(((singles - 4.09) / 0.833) * wpFactor * miscFactor)
                    singleCardNumbers = CInt(((singles - 3.68) / 0.833) * wpFactor * miscFactor)
                    'kCardNumbers = CInt(((strikeouts - 1.41) / 0.833) * wpFactor * miscFactor)
                    kCardNumbers = CInt(((strikeouts - 2.82) / 0.833) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt(((walks - 1.13) / 0.833) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt(((walks - 0.41) / 0.833) * wpFactor * miscFactor)
                    bbCardNumbers = CInt(((walks - 0.45) / 0.833) * wpFactor * miscFactor)
                Case "2-5"
                    'singleCardNumbers = CInt(((singles - 4.48) / 0.556) * wpFactor * miscFactor)
                    'singleCardNumbers = CInt(((singles - 8.46) / 0.556) * wpFactor * miscFactor)
                    singleCardNumbers = CInt(((singles - 7.89) / 0.556) * wpFactor * miscFactor)
                    'kCardNumbers = CInt(((strikeouts - 3.76) / 0.556) * wpFactor * miscFactor)
                    kCardNumbers = CInt(((strikeouts - 4.63) / 0.556) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt(((walks - 3) / 0.556) * wpFactor * miscFactor)
                    'bbCardNumbers = CInt(((walks - 1.74) / 0.556) * wpFactor * miscFactor)
                    bbCardNumbers = CInt(((walks - 1.18) / 0.556) * wpFactor * miscFactor)
                Case Else
                    Call MsgBox("PRB not set", vbOKOnly, "DeterminePitcherNumbers")
            End Select

            If singleCardNumbers < 0 Then singleCardNumbers = 0
            If kCardNumbers < 0 Then kCardNumbers = 0
            If bbCardNumbers < 0 Then bbCardNumbers = 0

        Catch ex As Exception
            Call MsgBox("DeterminePitcherNumbers " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' determines the wild pitch rating of a pitcher
    ''' </summary>
    ''' <param name="totalNumbers">The total numbers on the pitcher card. Used to handle
    ''' situations where the card numbers will exceed 88</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DetermineWPNumbers(ByVal totalNumbers As Integer) As Integer
        Dim outcomeRatio As Double
        Dim cardNumbers As Double
        Dim miscFactor As Double = 64 / totalNumbers

        Try
            outcomeRatio = (_BFP - _IBB) / 128
            'The * 5 takes into account that there are baserunners only half the time (*2, inverse of half),
            'and that there is only a 40% chance of a WP once it is chosen by the FAC card,
            '(*2.5, because it is the inverse of 40%). 2.5*2 is 5
            cardNumbers = (_WP / outcomeRatio) * 2.5
            'Adjust for PB rating, since a 2-9 pitcher will get many more references to their card
            'than a 2-5
            Select Case _PBR
                Case "2-9"
                    cardNumbers = cardNumbers * 3 / 5
                Case "2-8"
                    cardNumbers = cardNumbers * 9 / 13
                Case "2-7"
                    cardNumbers = cardNumbers * 6 / 7
                Case "2-6"
                    cardNumbers = cardNumbers * 6 / 5
                Case "2-5"
                    cardNumbers = cardNumbers * 9 / 5
            End Select
            cardNumbers *= miscFactor
            If cardNumbers > 3 Then cardNumbers = 3
        Catch ex As Exception
            Call MsgBox("DetermineWPNumbers " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return CInt(cardNumbers)
    End Function

    ''' <summary>
    ''' determines balk rating of pitcher
    ''' </summary>
    ''' <param name="totalNumbers">The total numbers on the pitcher card. Used to handle
    ''' situations where the card numbers will exceed 88</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DetermineBKNumbers(ByVal totalNumbers As Integer) As Integer
        Dim outcomeRatio As Double
        Dim cardNumbers As Double
        Dim miscFactor As Double = 64 / totalNumbers

        Try
            outcomeRatio = (_BFP - _IBB) / 128
            'The * 5 takes into account that there are baserunners only half the time (*2, inverse of half),
            'and that there is only a 40% chance of a WP once it is chosen by the FAC card,
            '(*2.5, because it is the inverse of 40%). 2.5*2 is 5
            cardNumbers = (_BK / outcomeRatio) * 2.5
            'Adjust for PB rating, since a 2-9 pitcher will get many more references to their card
            'than a 2-5
            Select Case _PBR
                Case "2-9"
                    cardNumbers = cardNumbers * 3 / 5
                Case "2-8"
                    cardNumbers = cardNumbers * 9 / 13
                Case "2-7"
                    cardNumbers = cardNumbers * 6 / 7
                Case "2-6"
                    cardNumbers = cardNumbers * 6 / 5
                Case "2-5"
                    cardNumbers = cardNumbers * 9 / 5
            End Select
            cardNumbers *= miscFactor
            If cardNumbers > 3 Then cardNumbers = 3
        Catch ex As Exception
            Call MsgBox("DetermineWPNumbers " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return CInt(cardNumbers)
    End Function

    ''' <summary>
    ''' determines a the pitcher PB rating based on where they rank in their league in ERA
    ''' </summary>
    ''' <param name="cumulativeBFP">the pitchers are sorted by era. This varible represents the number of outs (or IP/3)
    ''' we have gone through since starting at the top of the order (by pitcher era). So basically the 2-9 pitchers will be
    ''' determined first, then the 2-8s, etc.</param>
    ''' <param name="playerIpOuts"></param>
    ''' <remarks></remarks>
    Public Sub GetPBRating(ByVal cumulativeBFP As Integer, ByVal playerIpOuts As Integer)
        Dim ipTest As Double
        Dim pbRating As String

        Try
            If playerIpOuts > 0 Then
                ipTest = cumulativeBFP + CDbl(Format(playerIpOuts * 0.5, "###.##"))
            End If
            Select Case ipTest
                Case 0 To _28
                    pbRating = "2-9"
                Case (_28 - 0.01) To _27
                    pbRating = "2-8"
                Case (_27 - 0.01) To _26
                    pbRating = "2-7"
                Case (_26 - 0.01) To _25
                    pbRating = "2-6"
                Case Else
                    pbRating = "2-5"
            End Select
            _PBR = pbRating
        Catch ex As Exception
            Call MsgBox("GetPBRating " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

End Class

Public Class FACView
    Private _action As String
    Private _fname As String
    Private _fvalue As String

    Public Property action() As String
        Get
            Return _action
        End Get
        Set(ByVal value As String)
            _action = value
        End Set
    End Property
    Public Property facField() As String
        Get
            Return _fname
        End Get
        Set(ByVal value As String)
            _fname = value
        End Set
    End Property
    Public Property facValue() As String
        Get
            Return _fvalue
        End Get
        Set(ByVal value As String)
            _fvalue = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal action As String, ByVal facField As String, ByVal facValue As String)
        _action = action
        _fname = facField
        _fvalue = facValue
    End Sub
End Class

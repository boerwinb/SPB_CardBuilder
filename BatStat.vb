Option Explicit On
Option Strict On

Public Class BatStat
    Private _player As String
    Private _team As String
    Private _games As Integer
    Private _ab As Integer
    Private _Runs As Integer
    Private _Hits As Integer
    Private _Doubles As Integer
    Private _Triples As Integer
    Private _HRs As Integer
    Private _BB As Integer
    Private _IBB As Integer
    Private _K As Integer
    Private _SB As Integer
    Private _SH As Integer
    Private _SF As Integer
    Private _HB As Integer
    Private _Bats As String
    Private _PO As Integer
    Private _Err As Integer
    Private _Assists As Integer
    Private _DPS As Integer
    Private _PosGames As Integer
    Private _CurPos As String
    Private _CurTeam As String
    Private _adjSingle As Double
    Private _adjK As Double
    Private _adjBB As Double

    Public Sub Clear()
        _player = ""
        _team = ""
        _games = 0
        _ab = 0
        _Runs = 0
        _Hits = 0
        _Doubles = 0
        _Triples = 0
        _HRs = 0
        _BB = 0
        _IBB = 0
        _K = 0
        _SB = 0
        _SH = 0
        _SF = 0
        _HB = 0
        _Bats = ""
        _PO = 0
        _Err = 0
        _Assists = 0
        _DPS = 0
        _PosGames = 0
        _CurPos = ""
        _CurTeam = ""
        _adjSingle = 0
        _adjK = 0
        _adjBB = 0
    End Sub

    Public Property team() As String
        Get
            team = _team
        End Get
        Set(ByVal value As String)
            _team = value
        End Set
    End Property

    Public Property games() As Integer
        Get
            games = _games
        End Get
        Set(ByVal value As Integer)
            _games = value
        End Set
    End Property

    Public Property ab() As Integer
        Get
            ab = _ab
        End Get
        Set(ByVal value As Integer)
            _ab = value
        End Set
    End Property

    Public Property runs() As Integer
        Get
            runs = _Runs
        End Get
        Set(ByVal value As Integer)
            _Runs = value
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

    Public Property doubles() As Integer
        Get
            doubles = _Doubles
        End Get
        Set(ByVal value As Integer)
            _Doubles = value
        End Set
    End Property

    Public Property triples() As Integer
        Get
            triples = _Triples
        End Get
        Set(ByVal value As Integer)
            _Triples = value
        End Set
    End Property

    Public Property hrs() As Integer
        Get
            hrs = _HRs
        End Get
        Set(ByVal value As Integer)
            _HRs = value
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

    Public Property sb() As Integer
        Get
            sb = _SB
        End Get
        Set(ByVal value As Integer)
            _SB = value
        End Set
    End Property

    Public Property sh() As Integer
        Get
            sh = _SH
        End Get
        Set(ByVal value As Integer)
            _SH = value
        End Set
    End Property

    Public Property sf() As Integer
        Get
            sf = _sf
        End Get
        Set(ByVal value As Integer)
            _sf = value
        End Set
    End Property

    Public Property hb() As Integer
        Get
            hb = _HB
        End Get
        Set(ByVal value As Integer)
            _HB = value
        End Set
    End Property

    Public Property bats() As String
        Get
            bats = _Bats
        End Get
        Set(ByVal value As String)
            _Bats = value
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

    Public Property posGames() As Integer
        Get
            posGames = _PosGames
        End Get
        Set(ByVal value As Integer)
            _PosGames = value
        End Set
    End Property

    Public Property curPos() As String
        Get
            curPos = _CurPos
        End Get
        Set(ByVal value As String)
            _CurPos = value
        End Set
    End Property

    Public Property curTeam() As String
        Get
            curTeam = _CurTeam
        End Get
        Set(ByVal value As String)
            _CurTeam = value
        End Set
    End Property

    Public Property player() As String
        Get
            player = _player
        End Get
        Set(ByVal Value As String)
            _player = Value
        End Set
    End Property

    Public Property adjSingle() As Double
        Get
            adjSingle = _adjSingle
        End Get
        Set(ByVal Value As Double)
            _adjSingle = Value
        End Set
    End Property

    Public Property adjK() As Double
        Get
            adjK = _adjK
        End Get
        Set(ByVal Value As Double)
            _adjK = Value
        End Set
    End Property

    Public Property adjBB() As Double
        Get
            adjBB = _adjBB
        End Get
        Set(ByVal Value As Double)
            _adjBB = Value
        End Set
    End Property

    ''' <summary>
    ''' builds the field line for the batting card
    ''' </summary>
    ''' <param name="cdRating"></param>
    ''' <param name="games"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFieldingLine(ByRef cdRating As String, ByVal games As Integer) As String
        Dim fieldLine As String = ""
        Dim percentValue As Double
        Dim fieldErrorRating As Integer
        Dim fieldPosition As String
        Dim armRating As String

        Try
            If _PosGames > 1 Then

                fieldPosition = _CurPos
                'Calculate percentage BEFORE PROJECTING OVER 162!
                If (_PO + _Assists + _Err) = 0 Then
                    percentValue = 0
                Else
                    percentValue = Val(Format((_PO + _Assists) / (_PO + _Assists + _Err), "0.###"))
                End If
                If _PosGames > 99 * games / 162 Then
                    'Extrapolate assists, DPs over 162
                    _Assists = CInt(Val(_Assists) * 162 / Val(_PosGames))
                    _DPS = CInt(Val(_DPS) * 162 / Val(_PosGames))
                End If

                Select Case fieldPosition
                    Case "1B"
                        Select Case percentValue
                            Case 1
                                fieldErrorRating = 0
                            Case 0.995 To 0.999
                                fieldErrorRating = 1
                            Case 0.99 To 0.994
                                fieldErrorRating = 2
                            Case 0.985 To 0.989
                                fieldErrorRating = 3
                            Case 0.98 To 0.984
                                fieldErrorRating = 4
                            Case 0.975 To 0.979
                                fieldErrorRating = 5
                            Case 0.97 To 0.974
                                fieldErrorRating = 6
                            Case 0.965 To 0.969
                                fieldErrorRating = 7
                            Case 0.96 To 0.964
                                fieldErrorRating = 8
                            Case 0.955 To 0.959
                                fieldErrorRating = 9
                            Case Else
                                fieldErrorRating = 10
                        End Select
                        If Val(_PosGames) < 10 Then
                            If fieldErrorRating < 5 Then
                                fieldErrorRating = 5 'Punish the inexperienced, not reward
                            End If
                        End If
                        fieldLine = fieldPosition & "-" & _PosGames.ToString & " E" & _
                                        fieldErrorRating.ToString & Space(1)
                        If Val(_PosGames) > 80 * games / 162 Then
                            Select Case _Assists
                                Case 104 To 117
                                    cdRating = "1/1B"
                                Case Is > 117
                                    cdRating = "2/1B"
                            End Select
                        End If
                    Case "2B", "SS", "C"
                        Select Case percentValue
                            Case 1
                                fieldErrorRating = 0
                            Case 0.985 To 0.999
                                fieldErrorRating = 1
                            Case 0.975 To 0.984
                                fieldErrorRating = 2
                            Case 0.965 To 0.974
                                fieldErrorRating = 3
                            Case 0.955 To 0.964
                                fieldErrorRating = 4
                            Case 0.945 To 0.954
                                fieldErrorRating = 5
                            Case 0.935 To 0.944
                                fieldErrorRating = 6
                            Case 0.925 To 0.934
                                fieldErrorRating = 7
                            Case 0.915 To 0.924
                                fieldErrorRating = 8
                            Case 0.905 To 0.914
                                fieldErrorRating = 9
                            Case Else
                                fieldErrorRating = 10
                        End Select
                        If Val(_PosGames) < 10 Then
                            If fieldErrorRating < 5 Then
                                fieldErrorRating = 5 'Punish the inexperienced, not reward
                            End If
                        End If
                        If fieldPosition = "C" Then
                            If Val(_PosGames) < 10 Then
                                armRating = "C"
                            ElseIf _Assists > 65 Then
                                armRating = "A"
                            Else
                                armRating = "B"
                            End If
                            fieldLine = fieldPosition & "-" & _PosGames.ToString & " E" & _
                                        fieldErrorRating.ToString & " T" & armRating & Space(1)
                        Else
                            fieldLine = fieldPosition & "-" & _PosGames.ToString & " E" & _
                                        fieldErrorRating.ToString & Space(1)
                        End If
                        If Val(_PosGames) > 80 * games / 162 Then
                            If fieldPosition = "C" Then
                                Select Case _DPS
                                    Case 9 To 10
                                        cdRating = "1/C"
                                    Case Is > 10
                                        cdRating = "2/C"
                                End Select
                            ElseIf fieldPosition = "2B" Then
                                Select Case _DPS
                                    Case 129 To 136
                                        cdRating = "1/2B"
                                    Case Is > 136
                                        cdRating = "2/2B"
                                End Select
                            Else
                                Select Case _DPS
                                    Case 97 To 105
                                        cdRating = "1/SS"
                                    Case Is > 105
                                        cdRating = "2/SS"
                                End Select
                            End If
                        End If
                    Case "3B"
                        Select Case percentValue
                            Case 1
                                fieldErrorRating = 0
                            Case 0.986 To 0.999
                                fieldErrorRating = 1
                            Case 0.976 To 0.985
                                fieldErrorRating = 2
                            Case 0.966 To 0.975
                                fieldErrorRating = 3
                            Case 0.956 To 0.965
                                fieldErrorRating = 4
                            Case 0.946 To 0.955
                                fieldErrorRating = 5
                            Case 0.936 To 0.945
                                fieldErrorRating = 6
                            Case 0.926 To 0.935
                                fieldErrorRating = 7
                            Case 0.916 To 0.925
                                fieldErrorRating = 8
                            Case 0.906 To 0.915
                                fieldErrorRating = 9
                            Case Else
                                fieldErrorRating = 10
                        End Select
                        If Val(_PosGames) < 10 Then
                            If fieldErrorRating < 5 Then
                                fieldErrorRating = 5 'Punish the inexperienced, not reward
                            End If
                        End If
                        fieldLine = fieldPosition & "-" & _PosGames.ToString & " E" & _
                                    fieldErrorRating.ToString & Space(1)
                        If Val(_PosGames) > 80 * games / 162 Then
                            Select Case _DPS
                                Case 28 To 42
                                    cdRating = "1/3B"
                                Case Is > 44
                                    cdRating = "2/3B"
                            End Select
                        End If
                    Case "OF"
                        Select Case percentValue
                            Case 1
                                fieldErrorRating = 0
                            Case 0.99 To 0.999
                                fieldErrorRating = 1
                            Case 0.98 To 0.989
                                fieldErrorRating = 2
                            Case 0.97 To 0.979
                                fieldErrorRating = 3
                            Case 0.96 To 0.969
                                fieldErrorRating = 4
                            Case 0.95 To 0.959
                                fieldErrorRating = 5
                            Case 0.94 To 0.949
                                fieldErrorRating = 6
                            Case 0.93 To 0.939
                                fieldErrorRating = 7
                            Case 0.92 To 0.929
                                fieldErrorRating = 8
                            Case 0.91 To 0.919
                                fieldErrorRating = 9
                            Case Else
                                fieldErrorRating = 10
                        End Select
                        If _PosGames < 10 Then
                            If fieldErrorRating < 5 Then
                                fieldErrorRating = 5 'Punish the inexperienced, not reward
                            End If
                            armRating = "2"
                        Else
                            If Val(_PosGames) > 80 * games / 162 Then
                                Select Case _Assists
                                    Case 0 To 7
                                        armRating = "3"
                                    Case 8 To 9
                                        armRating = "3"
                                        cdRating = "1/OF"
                                    Case 10 To 11
                                        armRating = "4"
                                        cdRating = "1/OF"
                                    Case 12 To 13
                                        armRating = "4"
                                        cdRating = "2/OF"
                                    Case Else
                                        armRating = "5"
                                        cdRating = "2/OF"
                                End Select
                            Else
                                armRating = "3"
                            End If
                        End If
                        fieldLine = fieldPosition & "-" & _PosGames.ToString & " E" & _
                                fieldErrorRating.ToString & " T" & armRating & Space(1)
                    Case "DH"
                        fieldLine = fieldPosition & "-" & _PosGames.ToString & Space(1)
                End Select
            End If
        Catch ex As Exception
            Call MsgBox("GetFieldingLine " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return fieldLine
    End Function

    ''' <summary>
    ''' determines adjustment factors for singles, strikeouts, walks based on league averages for that season
    ''' </summary>
    ''' <param name="league"></param>
    ''' <param name="year"></param>
   ''' <remarks></remarks>
    Public Sub DetermineLeagueAvgs(ByVal league As String, ByVal year As Integer, ByRef DataAccess As clsDataAccess)
        Dim outcomeRatio As Double
        Dim totalSingles As Integer
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
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(.Item("totalbfp"))
                End With
            Else
                'Otherwise, try to determine from other stats
                sqlQuery = "SELECT sum(ipouts) as totalouts, sum(bb) as totalbb, sum(ibb) as totalibb, " & _
                                "sum(h) as totalhits FROM pitching WHERE lgid = '" & _
                            league & "' AND yearid = " & year.ToString
                'rs.Open(sqlQuery, cn)
                ds = DataAccess.ExecuteDataSet(sqlQuery)
                With ds.Tables(0).Rows(0)
                    totalBFP = CInt(CInt(.Item("totalouts")) * (26 / 27)) + CInt(.Item("totalbb")) + _
                                    CInt(.Item("totalhits")) ' - CheckField(.item("totalibb"), 0)
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

            totalSingles = totalH - totalD - totalT - totalHR
            outcomeRatio = (totalBFP - totalIBBS - totalSH) / 128
            _adjSingle = (totalSingles / outcomeRatio) / 2
            _adjK = (totalK / outcomeRatio) / 2
            _adjBB = ((totalBB - totalIBBS) / outcomeRatio) / 2
        Catch ex As Exception
            Call MsgBox("DetermineLeagueAvgs " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class

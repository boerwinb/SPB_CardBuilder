'Imports MySql.Data.MySqlClient
Imports System.Data.SQLite

Public Class clsDataAccess
    'Private _connection As mysqlConnection
    Private _connectionLite As SQLiteConnection
    Private _databaseLite As String
    'Private _server As String
    'Private _database As String
    'Private _uid As String
    'Private _password As String


    Public Sub New(ByVal database As String)
        Initialize(database)
    End Sub

    Private Sub Initialize(ByVal database As String)
        'Dim connectionString As String = ""

        '_server = "localhost"
        '_database = database
        _databaseLite = String.Format("Data Source={0}\{1}.db3", GetAppLocation, database)
        _connectionLite = New SQLiteConnection(_databaseLite)
        '_uid = "root"
        '_password = ""
        'connectionString = "SERVER=" & _server & ";DATABASE=" & _database & ";UID=" & _uid & ";PASSWORD=" & _password & ";"
        '_connection = New MySqlConnection(connectionString)
    End Sub

    Private Function OpenConnection() As Boolean
        Try
            '_connection.Open()
            If _connectionLite.State <> ConnectionState.Open Then
                _connectionLite.Open()
            End If
            Return True
            'Catch ex As MySqlException
        Catch ex As SQLiteException
            'Select Case ex.Number
            '    Case 0
            '        MessageBox.Show("Cannot connect to server. Contact administrator.")
            '    Case 1045
            '        MessageBox.Show("Invalid username/password, please try again")
            'End Select
            MessageBox.Show(String.Format("Error Num: {0}, Error Desc: {1}", ex.ErrorCode, ex.Message))
            Return False
        End Try
    End Function

    Private Function CloseConnection() As Boolean
        Try
            '_connection.Close()
            _connectionLite.Close()
            Return True
            'Catch ex As MySqlException
        Catch ex As SQLiteException
            'MessageBox.Show(ex.Message)
            MessageBox.Show(String.Format("Error Num: {0}, Error Desc: {1}", ex.ErrorCode, ex.Message))
            Return False
        End Try
    End Function

    Public Function ExecuteDataSet(ByVal sqlQuery As String) As DataSet
        Try
            'Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
            '    accessConn.ConnectionString = connectString
            '    accessConn.Open()
            'Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand
            If Me.OpenConnection Then
                'Using mysqlCmd As MySqlCommand = _connection.CreateCommand
                Using sqlLiteCmd As SQLiteCommand = _connectionLite.CreateCommand
                    'Using mysqlAdapter As New MySqlDataAdapter(mysqlCmd)
                    Using sqlLiteAdapter As New SQLiteDataAdapter(sqlLiteCmd)
                        Using ds As New DataSet
                            'mysqlCmd.CommandText = sqlQuery
                            sqlLiteCmd.CommandText = sqlQuery
                            'mysqlAdapter.Fill(ds)
                            sqlLiteAdapter.Fill(ds)
                            Me.CloseConnection()
                            Return ds
                        End Using
                    End Using
                End Using
            End If
            'End Using

            'Catch ex As MySqlException
        Catch ex As SQLiteException
            'Call MsgBox("Error in ExecuteDataSet. " & ex.ToString, MsgBoxStyle.OkOnly)
            MessageBox.Show(String.Format("Error Num: {0}, Error Desc: {1}", ex.ErrorCode, ex.Message))
        End Try
        Return Nothing
    End Function

    Public Sub Insert()

    End Sub

    Public Sub Update()

    End Sub

    Public Sub Delete()

    End Sub

    'Public Function SelectStatement() As List(Of String)

    'End Function

    Public Function Count() As Integer

    End Function

    Public Sub Backup()

    End Sub

    Public Sub Restore()

    End Sub

    ''' <summary>
    ''' determines how many games a team player during a given season
    ''' </summary>
    ''' <param name="colTeamGames"></param>
    ''' <param name="year"></param>
    ''' <remarks></remarks>
    Public Sub GetTeamGames(ByRef colTeamGames As Collection, ByVal year As String)
        'Dim rs As ADODB.Recordset
        Dim ds As DataSet = Nothing
        Dim sqlQuery As String
        Dim totalGames As Integer

        Try
            sqlQuery = "SELECT teamid, w, l from TEAMS WHERE yearid = " & year
            'If Me.OpenConnection() Then
            ds = ExecuteDataSet(sqlQuery)
            'While Not rs.EOF
            For Each dr As DataRow In ds.Tables(0).Rows
                totalGames = CInt(dr.Item("w")) + CInt(dr.Item("l"))
                colTeamGames.Add(totalGames, dr.Item("teamid").ToString)
            Next dr
            'Me.CloseConnection()
            'End If
        Catch ex As Exception
            Call MsgBox("GetTeamGames " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' returns full team name based on Lahman short name
    ''' </summary>
    ''' <param name="shortName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLongTeamTranslation(ByVal shortName As String) As String
        Dim dsTeam As DataSet = Nothing
        Dim sqlQuery As String = ""
        Dim longTeam As String = ""

        Try
            sqlQuery = "SELECT longteam FROM teamtranslation WHERE shortteam = '" & shortName & "'"
            'If Me.OpenConnection Then
            dsTeam = ExecuteDataSet(sqlQuery)
            With dsTeam.Tables(0).Rows(0)
                If dsTeam.Tables(0).Rows.Count = 0 Then
                    MsgBox("shortname not found!")
                Else
                    longTeam = .Item("longteam").ToString
                End If
            End With
            'Me.CloseConnection()
            'End If

        Catch ex As Exception
            Call MsgBox("Error in GetLongTeamTranslation. " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return longTeam
    End Function
End Class

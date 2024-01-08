Imports System.text

Module modRoutines
    Public gCurrentNum As Integer

    ''' <summary>
    ''' BB - overrides the traditional IIF method so that any data type can be returned
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Expression"></param>
    ''' <param name="TruePart"></param>
    ''' <param name="FalsePart"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IIF(Of T)(ByVal Expression As Boolean, ByVal TruePart As T, ByVal FalsePart As T) As T
        If Expression Then Return TruePart Else Return FalsePart
    End Function

    ''' <summary>
    ''' Builds Access connection string for Lahman tables
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLahmanConnectString() As String
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetAppLocation() & "\lahman56.mdb;User Id=admin;Password=;"
    End Function

    ''' <summary>
    ''' Builds Access connection string for FAC tables
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFACConnectString() As String
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetAppLocation() & "\fac.mdb;User Id=admin;Password=;"
    End Function

    ''' <summary>
    ''' Builds WHERE IN clause to identify pitchers within the batting table
    ''' </summary>
    ''' <param name="colPitchers"></param>
    ''' <returns></returns>
    ''' <remarks>BB Created:12/20/2008</remarks>
    Public Function BuildPitcherList(ByVal colPitchers As Collection) As String
        Dim sb As New stringbuilder

        For Each pitcherId As String In colPitchers
            sb.Append(",'")
            sb.Append(pitcherId)
            sb.Append("'")
        Next
        If colPitchers.Count > 0 Then
            'eliminate lead comma
            Return sb.ToString.Substring(1)
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' determines the next number to go on the card
    ''' </summary>
    ''' <param name="totalNumbers"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNumber(ByVal totalNumbers As Integer) As String
        Dim firstNumber As Integer
        Dim lastNumber As Integer
        Dim finalRange As String = ""

        Try
            If totalNumbers = 0 Or gCurrentNum = 88 Then
                Return finalRange
            End If
            If gCurrentNum = 0 Then
                firstNumber = 11
            ElseIf gCurrentNum Mod 10 = 8 Then
                'end of a base 8 series
                firstNumber = gCurrentNum + 3
            Else
                firstNumber = gCurrentNum + 1
            End If
            If totalNumbers = 1 Or firstNumber = 88 Then
                gCurrentNum = firstNumber
                finalRange = firstNumber.ToString
            Else
                lastNumber = firstNumber
                For i As Integer = 2 To totalNumbers
                    If lastNumber < 88 Then
                        If lastNumber Mod 10 = 8 Then
                            lastNumber = lastNumber + 3
                        Else
                            lastNumber = lastNumber + 1
                        End If
                    End If
                Next i
                gCurrentNum = lastNumber
                finalRange = firstNumber.ToString & "-" & lastNumber.ToString
            End If
        Catch ex As Exception
            Call MsgBox("GetNumber " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
        Return finalRange
    End Function

    ''' <summary>
    ''' Distribute hits according to the side of plate they swing from. Algorithm accuentuates left field for righties and 
    ''' right field for lefties
    ''' </summary>
    ''' <param name="totatHitNumbers"></param>
    ''' <param name="b7Hits"></param>
    ''' <param name="b8Hits"></param>
    ''' <param name="b9Hits"></param>
    ''' <param name="battingSide"></param>
    ''' <remarks></remarks>
    Public Sub AssignHitFields(ByVal totatHitNumbers As Integer, ByRef b7Hits As Integer, ByRef b8Hits As Integer, _
                        ByRef b9Hits As Integer, ByVal battingSide As String)
        Try
            For i As Integer = 1 To totatHitNumbers
                Select Case battingSide
                    Case "L"
                        If b9Hits = b8Hits And b9Hits - b7Hits < 2 Then
                            b9Hits += 1
                        ElseIf b7Hits = b8Hits Then
                            b8Hits += 1
                        Else
                            b7Hits += 1
                        End If
                    Case "S"
                        If b7Hits = b8Hits And b8Hits = b9Hits Then
                            b8Hits += 1
                        ElseIf b7Hits = b9Hits Then
                            b7Hits += 1
                        Else
                            b9Hits += 1
                        End If
                    Case "R"
                        If b7Hits = b8Hits And b7Hits - b9Hits < 2 Then
                            b7Hits += 1
                        ElseIf b9Hits = b8Hits Then
                            b8Hits += 1
                        Else
                            b9Hits += 1
                        End If
                End Select
            Next i
        Catch ex As Exception
            Call MsgBox("AssignHitFields " & ex.ToString, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Public Function GetAppLocation() As String
        Dim appPath As String
        Dim stringPosition As Integer

        appPath = Reflection.Assembly.GetExecutingAssembly.Location
        stringPosition = InStrRev(appPath, "\")
        If stringPosition > 0 Then
            appPath = appPath.Substring(0, stringPosition - 1)
        End If
        Return appPath
    End Function

    Public Function CheckField(ByVal value As Object, ByVal defaultValue As String) As String
        If IsDBNull(value) Then
            Return defaultValue
        Else
            Return value.ToString
        End If
    End Function

    Public Function CheckField(ByVal value As Object, ByVal defaultValue As Integer) As Integer
        If IsDBNull(value) Then
            Return defaultValue
        Else
            Return CInt(value)
        End If
    End Function

    'Public Function ExecuteDataSet(ByVal sqlQuery As String, ByVal connectString As String) As DataSet
    '    Try
    '        Using accessConn As OleDb.OleDbConnection = New OleDb.OleDbConnection
    '            accessConn.ConnectionString = connectString
    '            accessConn.Open()
    '            Using accessCmd As OleDb.OleDbCommand = accessConn.CreateCommand
    '                Using accessAdapter As New OleDb.OleDbDataAdapter(accessCmd)
    '                    Using ds As New DataSet
    '                        accessCmd.CommandText = sqlQuery
    '                        accessAdapter.Fill(ds)
    '                        Return ds
    '                    End Using
    '                End Using
    '            End Using
    '        End Using

    '    Catch ex As Exception
    '        Call MsgBox("Error in ExecuteDataSet. " & ex.ToString, MsgBoxStyle.OkOnly)
    '    End Try
    '    Return Nothing
    'End Function
End Module

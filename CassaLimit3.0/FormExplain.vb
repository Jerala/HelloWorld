Public Class FormExplain

    Public TypeInt As Integer

    Public SqlStr As String

    Public DT As Date
    Public DT2 As Date

    Public Cu As String
    Public P As String
    Public F As String
    Public NeedPodr As Boolean



    Private Sub FormExplain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dt_tbl As New DataTable

        Dim conn As New OleDb.OleDbConnection
        Dim com As New OleDb.OleDbCommand
        Dim Adap As New OleDb.OleDbDataAdapter
        Dim all As Boolean = False
        conn.ConnectionString = ConnectToAccess(all)


        With com
            .CommandType = CommandType.Text
            .Connection = conn
            .CommandText = SqlStr
        End With



        If all = True Then

            If TypeInt = 1 Then
                With com

                    DT = DT.ToShortDateString
                    .Parameters.AddWithValue("@Cur", Cu)
                    .Parameters.AddWithValue("@Dtt", DT)
                    If NeedPodr = True Then
                        .Parameters.AddWithValue("@P", P)
                    End If
                    .Parameters.AddWithValue("@F", F)
                    NeedPodr = False
                End With

            ElseIf TypeInt = 2 Then
                With com
                    DT = DT.ToShortDateString
                    DT2 = DT2.ToShortDateString
                    .Parameters.AddWithValue("@Cur", Cu)
                    .Parameters.AddWithValue("@Dtt", DT)
                    .Parameters.AddWithValue("@Dtt2", DT2)
                End With

            ElseIf TypeInt = 3 Then
                With com
                    DT = DT.ToShortDateString
                    DT2 = DT2.ToShortDateString
                    .Parameters.AddWithValue("@Cur", Cu)
                    .Parameters.AddWithValue("@Dtt", DT)
                    .Parameters.AddWithValue("@Dtt2", DT2)
                End With
            End If


            Adap.SelectCommand = com

            Try

                Adap.Fill(dt_tbl)
                DataGridView1.DataSource = dt_tbl
                MarofetDataGreed(DataGridView1, dt_tbl)
            Catch ex As Exception

                MsgBox(ex.Message)

            Finally
                conn.Dispose()
                com.Dispose()
                Adap.Dispose()
            End Try

        End If


    End Sub
    Public Sub MarofetDataGreed(dgv As DataGridView, dttbl As DataTable)

        Dim s As String, TypeColum As String

        For i As Integer = 0 To dttbl.Columns.Count - 1
            TypeColum = dttbl.Columns(i).DataType.Name
            s = Replace(dttbl.Columns(i).ColumnName, ", ", " ")

            If TypeColum = "Double" Or TypeColum = "Int32" Then

                With dgv.Columns(s).DefaultCellStyle

                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.BottomRight

                End With

            ElseIf TypeColum = "DateTime" Then

                dgv.Columns(s).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            ElseIf TypeColum = "String" Then

                dgv.Columns(s).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

            End If

        Next

    End Sub

End Class
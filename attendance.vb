Imports System.Data.DataTable
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Public Class Employee_Attendance
  Dim connection As SqlConnection = New SqlConnection("Data Source=Enter connection string here")


    Private Sub Employee_Attendance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PayaldbDataSet8.emp_reg' table. You can move, or remove it, as needed.
        'Me.Emp_regTableAdapter.Fill(Me.PayaldbDataSet8.emp_reg)

    End Sub

    Private Sub btn_open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_open.Click
        'query is to fetch data emp_name from sql  
        connection.Open()

        ' Check if the data for that selected date is present in emp_attendance table or not present

        Dim cmd As New SqlCommand("select count(*) from emp_attendance where Date=@Date", connection)
        Dim adptor As New SqlDataAdapter(cmd)
        cmd.Parameters.Add(New SqlParameter With {.ParameterName = "@Date", .SqlDbType = SqlDbType.Date, .Value = DateTimePicker1.Value})
        Dim DT As New DataTable
        adptor.Fill(DT)
        DataGridView1.Columns.Clear()

        If DT.Rows(0)(0) <> 0 Then
            DataGridView1.ReadOnly = True
            ' If data is present show data to user from emp_attendance
            Dim cmd2 As New SqlCommand("select Id, Emp_Name, Status from emp_attendance where Date=@Date", connection)
            cmd2.Parameters.Add(New SqlParameter With {.ParameterName = "@Date", .SqlDbType = SqlDbType.Date, .Value = DateTimePicker1.Value})
            Dim dbAdopter As New SqlDataAdapter(cmd2)
            Dim dataSet As New DataSet()
            dbAdopter.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)
            btn_submit.Enabled = False

        Else
            ' If data is not present show data to user from emp_reg
            DataGridView1.ReadOnly = False
            Dim cmd2 As New SqlCommand("select Id, Emp_Name from emp_reg", connection)
            Dim dbAdopter As New SqlDataAdapter(cmd2)
            Dim dataSet As New DataSet()
            dbAdopter.Fill(dataSet)
            DataGridView1.ClearSelection()
            DataGridView1.DataSource = dataSet.Tables(0)

            Dim checkBoxColumnPresent As New DataGridViewCheckBoxColumn()
            checkBoxColumnPresent.HeaderText = "Present"
            checkBoxColumnPresent.Width = 30
            checkBoxColumnPresent.Name = "Present"
            checkBoxColumnPresent.TrueValue = True
            DataGridView1.Columns.Insert(2, checkBoxColumnPresent)

            Dim checkBoxColumnAbsent As New DataGridViewCheckBoxColumn()
            checkBoxColumnAbsent.HeaderText = "Absent"
            checkBoxColumnAbsent.Width = 30
            checkBoxColumnAbsent.Name = "Absent"
            DataGridView1.Columns.Insert(3, checkBoxColumnAbsent)
            Dim checkBoxColumnLeave As New DataGridViewCheckBoxColumn()
            checkBoxColumnLeave.HeaderText = "Leave"
            checkBoxColumnLeave.Width = 30
            checkBoxColumnLeave.Name = "Leave"
            DataGridView1.Columns.Insert(4, checkBoxColumnLeave)
            btn_submit.Enabled = True
            DataGridView1.AllowUserToAddRows = False
            If DataGridView1.RowCount <> 0 Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    'Dim chkbox As DataGridViewCheckBoxCell = TryCast(row.Cells("Present"), DataGridViewCheckBoxCell)

                    'chkbox.Value = headerChk.AutoCheck

                    DataGridView1.Item(2, row.Index).Value = True
                Next
            End If
        End If
 
        connection.Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick


    End Sub

    Private Sub btn_submit_Click(sender As System.Object, e As System.EventArgs) Handles btn_submit.Click
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim present As Boolean = CBool(row.Cells("Present").Value)
            Dim absent As Boolean = CBool(row.Cells("Absent").Value)
            Dim leave As Boolean = CBool(row.Cells("Leave").Value)
            If present = False And absent = False And leave = False Then
                MessageBox.Show("Please select valid value")
                Exit Sub

            End If
        Next

        connection.Open()
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim present As Boolean = CBool(row.Cells("Present").Value)
            Dim absent As Boolean = CBool(row.Cells("Absent").Value)
            Dim leave As Boolean = CBool(row.Cells("Leave").Value)
            Dim cmd2 As New SqlCommand("insert into emp_attendance(Id,Emp_Name,Date,Status) values(@Id,@Emp_Name,@Date,@Status)", connection)
            cmd2.Parameters.AddWithValue("Id", row.Cells("Id").Value)
            cmd2.Parameters.AddWithValue("Emp_Name", row.Cells("Emp_Name").Value)


            If present Then
                cmd2.Parameters.AddWithValue("Status", "Present")
            End If
            If absent Then
                cmd2.Parameters.AddWithValue("Status", "Absent")
            End If
            If leave Then
                cmd2.Parameters.AddWithValue("Status", "Leave")
            End If
            cmd2.Parameters.Add(New SqlParameter With {.ParameterName = "@Date", .SqlDbType = SqlDbType.Date, .Value = DateTimePicker1.Value})

            cmd2.ExecuteNonQuery()
        Next
        connection.Close()
        MessageBox.Show("Data inserted Successfully")
    End Sub
    Private Sub Panel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub DataGridView1_CellValidated(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' Column Index =2 means Present
        ' Column Index =3 means Absent
        ' Column Index =4 means Leave
        ' DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value means value of current cell which is changing

        If DataGridView1.ReadOnly = False & e.ColumnIndex > 2
            If e.ColumnIndex = 2 And DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = True Then
                DataGridView1.Item(3, e.RowIndex).Value = False
                DataGridView1.Item(4, e.RowIndex).Value = False
            End If
            If e.ColumnIndex = 3 And DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = True Then
                DataGridView1.Item(2, e.RowIndex).Value = False
                DataGridView1.Item(4, e.RowIndex).Value = False

            End If
            If e.ColumnIndex = 4 And DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = True Then
                DataGridView1.Item(3, e.RowIndex).Value = False
                DataGridView1.Item(2, e.RowIndex).Value = False
            End If
        End If

    End Sub
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As System.EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
End Class

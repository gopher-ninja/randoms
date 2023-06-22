Imports System.Data.DataTable
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Public Class Employee_Attendance
    Dim connection As SqlConnection = New SqlConnection("Data Source=ROYAL\MSSQLSERVER04;Initial Catalog=payaldb;Integrated Security=True")


    Private Sub Employee_Attendance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PayaldbDataSet8.emp_reg' table. You can move, or remove it, as needed.
        'Me.Emp_regTableAdapter.Fill(Me.PayaldbDataSet8.emp_reg)

    End Sub

    Private Sub btn_open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_open.Click


        ' Step-1: Check data is available for SELECTED DATE in emp_attendance table
        ' Step-2: If data is present in emp_attendance table then just data from emp_attendance table for selected date
        ' Step-3: If data is NOT present in emp_attendance then perform Step-4
        ' Step-4: Get(Select statement) Emp_Name data from emp_reg table in that way we will get unique emp_Name data ( As in emp_attendance table Emp_Name is repeated)
        ' Step-5: Visualize data in datagridview with present default checked

        connection.Open()

        ' Check if the data for that selected date is present in emp_attendance table or not present

        Dim cmd As New SqlCommand("select Id, Emp_Name, Status from emp_attendance where Date=@Date", connection)
        Dim adptor As New SqlDataAdapter(cmd)
        cmd.Parameters.Add(New SqlParameter With {.ParameterName = "@Date", .SqlDbType = SqlDbType.Date, .Value = DateTimePicker1.Value})
        Dim DT As New DataTable
        adptor.Fill(DT)
        ' Clear old data present in datagridview1
        DataGridView1.Columns.Clear()

        ' Check if the row count is greater than 0 that means data is present for selected date in emp_attendance table
        If DT.Rows.Count > 0 Then

            ' As data is present in emp_attendance then we will show data from emp_attendance table for selected date
            ' TO Restrict user to edit any data set read only  = true
            DataGridView1.ReadOnly = True
            ' If data is present show data to user from emp_attendance
            ' data will get filled in dbAdapter
            ' Data will get visualize in DataGridView1 using following step (setting data source)
            DataGridView1.DataSource = DT.Rows
            btn_submit.Enabled = False

        Else
            ' Data is not present in emp_attendance so we will fetch data(Emp_Name) from emp_reg
            ' If data is not present show data to user from emp_reg
            DataGridView1.ReadOnly = False ' So that user can edit the data P A L
            Dim cmd2 As New SqlCommand("select Id, Emp_Name from emp_reg", connection)
            Dim dbAdopter As New SqlDataAdapter(cmd2)
            Dim dataSet As New DataSet()
            dbAdopter.Fill(dataSet)
            DataGridView1.ClearSelection()
            DataGridView1.DataSource = dataSet.Tables(0)

            ' Create a column for present
            Dim checkBoxColumnPresent As New DataGridViewCheckBoxColumn()
            checkBoxColumnPresent.HeaderText = "Present"
            checkBoxColumnPresent.Width = 30
            checkBoxColumnPresent.Name = "Present"
            checkBoxColumnPresent.TrueValue = True
            DataGridView1.Columns.Insert(2, checkBoxColumnPresent)

            ' Create a column for absent
            Dim checkBoxColumnAbsent As New DataGridViewCheckBoxColumn()
            checkBoxColumnAbsent.HeaderText = "Absent"
            checkBoxColumnAbsent.Width = 30
            checkBoxColumnAbsent.Name = "Absent"
            DataGridView1.Columns.Insert(3, checkBoxColumnAbsent)

            ' Create a column for Leave
            Dim checkBoxColumnLeave As New DataGridViewCheckBoxColumn()
            checkBoxColumnLeave.HeaderText = "Leave"
            checkBoxColumnLeave.Width = 30
            checkBoxColumnLeave.Name = "Leave"
            DataGridView1.Columns.Insert(4, checkBoxColumnLeave)

            ' Enable Submit button
            btn_submit.Enabled = True

            ' To restrict user from adding and deleting row
            DataGridView1.AllowUserToAddRows = False
            DataGridView1.AllowUserToDeleteRows = False

            'To make "Present" cell "Checked" we will have to check each rows of the data grid view
            If DataGridView1.RowCount > 0 Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    'To make "Present" cell "Checked" by setting its value True
                    DataGridView1.Item(2, row.Index).Value = True
                Next
            End If
        End If

        connection.Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick


    End Sub

    Private Sub btn_submit_Click(sender As System.Object, e As System.EventArgs) Handles btn_submit.Click
        ' To make sure atleast 1 of P, A or L should checked
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim present As Boolean = CBool(row.Cells("Present").Value)    ' Convert present cell value into boolean
            Dim absent As Boolean = CBool(row.Cells("Absent").Value) ' Convert absent cell value into boolean
            Dim leave As Boolean = CBool(row.Cells("Leave").Value) ' Convert leave cell value into boolean
            ' If all are unchecked then give warning to user
            If present = False And absent = False And leave = False Then
                MessageBox.Show("Please select at least one checkbox")
                Exit Sub

            End If
        Next

        connection.Open()
        ' Inserting data into table for each row after PROPER VALIDATION 
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim present As Boolean = CBool(row.Cells("Present").Value)
            Dim absent As Boolean = CBool(row.Cells("Absent").Value)
            Dim leave As Boolean = CBool(row.Cells("Leave").Value)
            Dim insertCmd As New SqlCommand("insert into emp_attendance(Id,Emp_Name,Date,Status) values(@Id,@Emp_Name,@Date,@Status)", connection)
      
            If present Then
                insertCmd.Parameters.AddWithValue("Status", "Present")
            End If
            If absent Then
                insertCmd.Parameters.AddWithValue("Status", "Absent")
            End If
            If leave Then
                insertCmd.Parameters.AddWithValue("Status", "Leave")
            End If
            ' To add value for query 
            insertCmd.Parameters.AddWithValue("Id", row.Cells("Id").Value)
            insertCmd.Parameters.AddWithValue("Emp_Name", row.Cells("Emp_Name").Value)
            insertCmd.Parameters.Add(New SqlParameter With {.ParameterName = "@Date", .SqlDbType = SqlDbType.Date, .Value = DateTimePicker1.Value})

            insertCmd.ExecuteNonQuery()
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

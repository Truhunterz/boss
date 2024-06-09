Form 1 :

Public Class Form1 

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Disable text fields and combo boxes by default
        txtID.Enabled = False
        txtName.Enabled = False
        txtIC.Enabled = False
        cbGender.Enabled = False
        txtNum.Enabled = False
        txtAdd.Enabled = False
        cbClinic.Enabled = False
        btnSave.Enabled = False
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' Enable text fields and combo boxes
        txtID.Enabled = True
        txtName.Enabled = True
        txtIC.Enabled = True
        cbGender.Enabled = True
        txtNum.Enabled = True
        txtAdd.Enabled = True
        cbClinic.Enabled = True
        btnSave.Enabled = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Check if any of the text boxes are empty
        If String.IsNullOrEmpty(txtID.Text) OrElse
           String.IsNullOrEmpty(txtName.Text) OrElse
           String.IsNullOrEmpty(txtIC.Text) OrElse
           cbGender.SelectedIndex < 0 OrElse
           String.IsNullOrEmpty(txtNum.Text) OrElse
           String.IsNullOrEmpty(txtAdd.Text) OrElse
           cbClinic.SelectedIndex < 0 Then

            MessageBox.Show("Please fill up all fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Try
                ' Open the connection
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If

                ' Insert command
                Dim insertQuery As String = "INSERT INTO Patient (PatID, Name, IC, Gender, Mobile, Address, Clinician) VALUES (@ID, @Name, @IC, @Gender, @PhoneNumber, @Address, @Clinic)"
                cmd = New OleDb.OleDbCommand(insertQuery, con)

                ' Add parameters
                cmd.Parameters.AddWithValue("@ID", txtID.Text)
                cmd.Parameters.AddWithValue("@Name", txtName.Text)
                cmd.Parameters.AddWithValue("@IC", txtIC.Text)
                cmd.Parameters.AddWithValue("@Gender", cbGender.Text)
                cmd.Parameters.AddWithValue("@PhoneNumber", txtNum.Text)
                cmd.Parameters.AddWithValue("@Address", txtAdd.Text)
                cmd.Parameters.AddWithValue("@Clinic", cbClinic.Text)

                ' Execute the command
                cmd.ExecuteNonQuery()

                MessageBox.Show("Form submitted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Clear the form fields after submission
                txtID.Clear()
                txtName.Clear()
                txtIC.Clear()
                cbGender.SelectedIndex = -1
                txtNum.Clear()
                txtAdd.Clear()
                cbClinic.SelectedIndex = -1

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                ' Close the connection
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        ' Confirm delete action
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this record?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            ' Check if the ID field is not empty
            If String.IsNullOrEmpty(txtID.Text) Then
                MessageBox.Show("Please enter the ID of the record to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Try
                    ' Open the connection
                    If con.State = ConnectionState.Closed Then
                        con.Open()
                    End If

                    ' Delete command
                    Dim deleteQuery As String = "DELETE FROM Patient WHERE PatID = ?"
                    cmd = New OleDb.OleDbCommand(deleteQuery, con)

                    ' Add parameter
                    cmd.Parameters.AddWithValue("?", txtID.Text)

                    ' Execute the command
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Record deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Clear the form fields after deletion
                        txtID.Clear()
                        txtName.Clear()
                        txtIC.Clear()
                        cbGender.SelectedIndex = -1
                        txtNum.Clear()
                        txtAdd.Clear()
                        cbClinic.SelectedIndex = -1

                        ' Disable text fields and combo boxes after deletion
                        txtID.Enabled = False
                        txtName.Enabled = False
                        txtIC.Enabled = False
                        cbGender.Enabled = False
                        txtNum.Enabled = False
                        txtAdd.Enabled = False
                        cbClinic.Enabled = False
                        btnSave.Enabled = False
                    Else
                        MessageBox.Show("No record found with the specified ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    ' Close the connection
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                End Try
            End If
        End If
    End Sub


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' Confirm exit action
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Exit the application
            Application.Exit()
        End If
    End Sub

    Private Sub cbClinician_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbClinic.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form2.Show()
    End Sub
End Class







module db:
Module Module1

    Public ds As New DataSet
    Public cmd As New OleDb.OleDbCommand
    Public da As New OleDb.OleDbDataAdapter
    Public con As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sfanz\OneDrive\Desktop\Databasevb\Patient.accdb")

End Module

























form 1/1:
Imports System.Data.OleDb

Public Class Form1
    ' Connection string to connect to Access database
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sfanz\OneDrive\Desktop\Databasevb\StudentDB.accdb;"

    ' Form Load event to initialize components
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Optional: Initialize components or data
    End Sub

    ' Event handler for the About menu item
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefToolStripMenuItem.Click
        MessageBox.Show("This is created by Syawal")
    End Sub

    ' Event handler for the Calculate button
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Try
            ' Retrieve form data
            Dim name As String = UCase(txtName.Text)
            Dim gender As String = cbGender.Text
            Dim phone As String = txtPhonenum.Text
            Dim course As String = cbCourse.Text
            Dim semester As Integer = Integer.Parse(txtSemester.Text)
            Dim fee As Decimal = Decimal.Parse(txtFee.Text)
            Dim total As Decimal = fee * semester

            ' Display data in a multiline TextBox
            txtShow.Text = $"Name : {name}{vbNewLine}Gender : {gender}{vbNewLine}Phone : {phone}{vbNewLine}Course : {course}{vbNewLine}Semester : {semester}{vbNewLine}Fee : RM{total}"
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
        MsgBox("Thank you.")
    End Sub

    ' Event handler for the Insert button
    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Try
            ' Retrieve form data
            Dim name As String = UCase(txtName.Text)
            Dim gender As String = cbGender.Text
            Dim phone As String = txtPhonenum.Text
            Dim course As String = cbCourse.Text
            Dim semester As Integer = Integer.Parse(txtSemester.Text)
            Dim fee As Decimal = Decimal.Parse(txtFee.Text)

            ' Insert data into the database
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "INSERT INTO Table1 (Name, Gender, Phone, Course, Semester, Fee) VALUES (?, ?, ?, ?, ?, ?)"
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@Name", name)
                    command.Parameters.AddWithValue("@Gender", gender)
                    command.Parameters.AddWithValue("@Phone", phone)
                    command.Parameters.AddWithValue("@Course", course)
                    command.Parameters.AddWithValue("@Semester", semester)
                    command.Parameters.AddWithValue("@Fee", fee)
                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Data inserted successfully.")
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    ' Event handler for the Close button
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' Method to load data from the database into a DataTable
    Private Function LoadData() As DataTable
        Dim dataTable As New DataTable()
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT * FROM Table1"
                Dim adapter As New OleDbDataAdapter(query, connection)
                adapter.Fill(dataTable)
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
        Return dataTable
    End Function

    ' Event handler to open Form2 and display the data in DataGridView
    Private Sub btnShowData_Click(sender As Object, e As EventArgs) Handles btnShowdata.Click
        Dim form2 As New Form2()
        form2.SetData(LoadData())
        form2.Show()
    End Sub
End Class

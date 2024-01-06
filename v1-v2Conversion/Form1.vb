Imports MySqlConnector

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadprodtableList()
        LoadDevTableList()

    End Sub
    Private Sub LoadprodtableList()
        Dim sqlData As String
        Dim connData As New MySqlConnection(GlobalVars.csProd)

        Try
            connData.Open()
            sqlData = "Select table_name FROM information_schema.tables WHERE table_schema = 'usalespl_TBE_Data';"
            Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
            Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
            If reader1Data.HasRows Then
                While reader1Data.Read()
                    ComboBox1.Items.Add(reader1Data.GetString(0))
                End While
            End If
            reader1Data.Close()
        Catch ex As Exception

        Finally
            connData.Close()
        End Try

    End Sub
    Private Sub LoadDevTableList()
        Dim sqlData As String
        Dim connData As New MySqlConnection(GlobalVars.csNewProd)

        Try
            connData.Open()
            sqlData = "Select  table_name FROM information_schema.tables WHERE table_schema = 'usalesplatform_usalesplatformTBEV2';"
            Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
            Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
            If reader1Data.HasRows Then
                While reader1Data.Read()
                    ComboBox2.Items.Add(reader1Data.GetString(0))
                End While
            End If
            reader1Data.Close()
        Catch ex As Exception

        Finally
            connData.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
    'rec orderdetail_id 
End Class

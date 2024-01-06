Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Security.Cryptography
Imports CefSharp
Imports Guna.Charts.WinForms
Imports Guna.UI2.WinForms
Imports MySql.Data.MySqlClient
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports MySqlConnector

Module GlobalRoutines

#Region "Example SQL Methods"
    'Public Sub ExampleUpdate()
    '    Dim sqlData As String
    '    Dim myCommand As New MySqlCommand
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        myCommand.Connection = connData
    '        sqlData = "UPDATE tbl_Example SET " &
    '                "Param1 = @Param1, " &
    '                "Param2 = @Param2 " &
    '                "WHERE Param3 = @Param3; "
    '        With myCommand
    '            .CommandType = CommandType.Text
    '            .CommandText = sqlData
    '            .Parameters.AddWithValue("@Param1", "")
    '            .Parameters.AddWithValue("@Param2", "")
    '            .Parameters.AddWithValue("@Param3", "")
    '        End With
    '        myCommand.ExecuteNonQuery()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub ExampleInsert()
    '    Dim sqlData1 As String
    '    Dim myCommand1 As New MySqlCommand
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        myCommand1.Connection = connData
    '        sqlData1 = "INSERT INTO tbl_Example (Param1, Param2, Param3) VALUES (@Param1, @Param2, @Param3);"
    '        With myCommand1
    '            .CommandText = sqlData1
    '            .Parameters.AddWithValue("@Param1", "")
    '            .Parameters.AddWithValue("@Param2", "")
    '            .Parameters.AddWithValue("@Param3", "")
    '        End With
    '        myCommand1.ExecuteNonQuery()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub ExampleSelect() ' A better select query
    '    Dim sqlData As String
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        sqlData = "SELECT something FROM tbl_Example WHERE Param1 = @Param1; "
    '        Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
    '        With cmdData
    '            .Parameters.AddWithValue("@Param1", "")
    '        End With
    '        Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
    '        If reader1Data.HasRows Then
    '            While reader1Data.Read()

    '            End While
    '        End If
    '        reader1Data.Close()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub ExampleDelete()
    '    Dim myCommand As New MySqlCommand
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        myCommand.Connection = connData
    '        Dim query As String = "DELETE FROM tbl_Example WHERE Param1 = @Param1;"
    '        With myCommand
    '            .CommandType = CommandType.Text
    '            .CommandText = query
    '            .Parameters.AddWithValue("@Param1", "")
    '        End With
    '        myCommand.ExecuteNonQuery()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub
#End Region

#Region "Test Subs"
    'Public NumRecsChanged As Integer = 0

    'Public Sub GetOrderDetailInfo()
    '    ' Dev Note: Use to get order info (grade, grdsvc, etc.) from tbl_OrderDetail
    '    Debug.WriteLine("--- Starting to retreive records...")

    '    Dim sqlData As String
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        sqlData = "SELECT RecID, Prod_ID, Grade, GrdSvc, Prod_Desc FROM tbl_OrderDetail WHERE Source IS NOT NULL AND Prod_RecID IS NULL"
    '        ' Dev Note: Check for updated amount
    '        'sqlData = "SELECT RecID, Prod_ID, Grade, GrdSvc, Prod_Desc FROM tbl_OrderDetail WHERE Source IS NOT NULL AND Prod_RecID IS NOT NULL"
    '        Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
    '        Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
    '        If reader1Data.HasRows Then
    '            While reader1Data.Read()
    '                Dim recID As Integer = 0
    '                Dim prodID As String = "None"
    '                Dim grade As String = "None"
    '                Dim grdsvc As String = "None"
    '                Dim desc As String = "None"

    '                If reader1Data.IsDBNull(0) = False Then
    '                    recID = reader1Data.GetValue(0)
    '                End If
    '                If reader1Data.IsDBNull(1) = False Then
    '                    prodID = reader1Data.GetString(1)
    '                End If
    '                If reader1Data.IsDBNull(2) = False Then
    '                    grade = reader1Data.GetString(2)
    '                End If
    '                If reader1Data.IsDBNull(3) = False Then
    '                    grdsvc = reader1Data.GetString(3)
    '                End If
    '                If reader1Data.IsDBNull(4) = False Then
    '                    desc = reader1Data.GetString(4)
    '                End If

    '                GetProdRecIDForOrderDetailFromProdsForSale(recID, prodID, grade, grdsvc, desc)
    '            End While
    '        End If
    '        reader1Data.Close()

    '        Debug.WriteLine("--- Number of Records Changed: " & NumRecsChanged)
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub GetProdRecIDForOrderDetailFromProdsForSale(recID As Integer, prodID As String, grade As String, grdSvc As String, desc As String)
    '    ' Dev Note: Use to get RecID from tbl_ProductForSale and save to tbl_OrderDetail as Prod_RecID
    '    Dim sqlData As String
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        sqlData = "SELECT RecID FROM tbl_ProductForSale WHERE Prod_ID = @Prod_ID AND (Grade = @Grade Or GrdSvc = @GrdSvc OR Prod_Desc = @Prod_Desc);"
    '        Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
    '        With cmdData
    '            .Parameters.AddWithValue("@Prod_ID", prodID)
    '            .Parameters.AddWithValue("@Grade", grade)
    '            .Parameters.AddWithValue("@GrdSvc", grdSvc)
    '            .Parameters.AddWithValue("@Prod_Desc", desc)
    '        End With
    '        Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
    '        If reader1Data.HasRows Then
    '            While reader1Data.Read()
    '                Dim prodrecID As Integer = reader1Data.GetValue(0)
    '                SetProdRecIDForOrderDetail(recID, prodrecID)
    '            End While
    '        End If
    '        reader1Data.Close()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub SetProdRecIDForOrderDetail(recID As Integer, prodRecID As Integer)
    '    Dim sqlData As String
    '    Dim myCommand As New MySqlCommand
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        myCommand.Connection = connData
    '        sqlData = "UPDATE tbl_OrderDetail SET " &
    '                "Prod_RecID = @Prod_RecID " &
    '                "WHERE RecID = @RecID; "
    '        With myCommand
    '            .CommandType = CommandType.Text
    '            .CommandText = sqlData
    '            .Parameters.AddWithValue("@Prod_RecID", prodRecID)
    '            .Parameters.AddWithValue("@RecID", recID)
    '        End With
    '        myCommand.ExecuteNonQuery()

    '        NumRecsChanged += 1
    '        'Debug.WriteLine("--- Setting Prod_RecID: " & prodRecID & " for OrderDetail RecID: " & recID)
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub GetOrderDetailsWithProdRecID()
    '    ' Dev Note: Use to get ProdRecID from tbl_OrderDetail
    '    Debug.WriteLine("--- Starting to retreive records...")
    '    Dim sqlData As String
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        sqlData = "SELECT RecID, Prod_RecID FROM tbl_OrderDetail WHERE Prod_RecID IS NOT NULL"
    '        Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
    '        Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
    '        If reader1Data.HasRows Then
    '            While reader1Data.Read()
    '                Dim recID As Integer = reader1Data.GetValue(0)
    '                Dim prodRecID As Integer = reader1Data.GetValue(1)
    '                GetProdsForSaleInfoWithProdRecID(recID, prodRecID)
    '            End While
    '        End If
    '        reader1Data.Close()

    '        Debug.WriteLine("--- Complete")
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub GetProdsForSaleInfoWithProdRecID(recID As Integer, prodRecID As Integer)
    '    ' Dev Note: Use to get ProdsForSale info using prodRecID

    '    Dim sqlData As String
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        sqlData = "SELECT Year, Denomination, MintMark, Grade, Prod_Desc, ExpiresOn, Cost, GrdSvc, Source, CoinTypeID, Metal, Weight, ShipWeight, Length, Width, Height, SpiffAmount, Tags FROM tbl_ProductForSale WHERE RecID = @RecID"
    '        Dim cmdData As MySqlCommand = New MySqlCommand(sqlData, connData)
    '        With cmdData
    '            .Parameters.AddWithValue("@RecID", prodRecID)
    '        End With
    '        Dim reader1Data As MySqlDataReader = cmdData.ExecuteReader()
    '        If reader1Data.HasRows Then
    '            While reader1Data.Read()
    '                Dim year As Integer = 0 '0
    '                Dim denomination As String = "" '1
    '                Dim mintMark As String = "" '2
    '                Dim grade As String = "" '3
    '                Dim desc As String = "" '4
    '                Dim expiresOn As Date
    '                Dim cost As Decimal = 0.0 '6
    '                Dim grdSvc As String = "" '7
    '                Dim source As String = "" '8
    '                Dim coinTypeID As Integer = 1 '9
    '                Dim metal As String = "" '10
    '                Dim weight As Decimal = 0.0 '11
    '                Dim shipWeight As Decimal = 0.0 '12
    '                Dim length As Decimal = 0.0 '13
    '                Dim width As Decimal = 0.0 '14
    '                Dim height As Decimal = 0.0 '15
    '                Dim spiffAmount As Decimal = 0.0 '16
    '                Dim tags As String = "" '17

    '                If reader1Data.IsDBNull(0) = False Then
    '                    year = reader1Data.GetValue(0)
    '                End If
    '                If reader1Data.IsDBNull(1) = False Then
    '                    denomination = reader1Data.GetString(1)
    '                End If
    '                If reader1Data.IsDBNull(2) = False Then
    '                    mintMark = reader1Data.GetString(2)
    '                End If
    '                If reader1Data.IsDBNull(3) = False Then
    '                    grade = reader1Data.GetString(3)
    '                End If
    '                If reader1Data.IsDBNull(4) = False Then
    '                    desc = reader1Data.GetString(4)
    '                End If
    '                If reader1Data.IsDBNull(5) = False Then
    '                    expiresOn = reader1Data.GetValue(5)
    '                End If
    '                If reader1Data.IsDBNull(6) = False Then
    '                    cost = reader1Data.GetValue(6)
    '                End If
    '                If reader1Data.IsDBNull(7) = False Then
    '                    grdSvc = reader1Data.GetString(7)
    '                End If
    '                If reader1Data.IsDBNull(8) = False Then
    '                    source = reader1Data.GetString(8)
    '                End If
    '                If reader1Data.IsDBNull(9) = False Then
    '                    coinTypeID = reader1Data.GetValue(9)
    '                End If
    '                If reader1Data.IsDBNull(10) = False Then
    '                    metal = reader1Data.GetString(10)
    '                End If
    '                If reader1Data.IsDBNull(11) = False Then
    '                    weight = reader1Data.GetValue(11)
    '                End If
    '                If reader1Data.IsDBNull(12) = False Then
    '                    shipWeight = reader1Data.GetValue(12)
    '                End If
    '                If reader1Data.IsDBNull(13) = False Then
    '                    length = reader1Data.GetValue(13)
    '                End If
    '                If reader1Data.IsDBNull(14) = False Then
    '                    width = reader1Data.GetValue(14)
    '                End If
    '                If reader1Data.IsDBNull(15) = False Then
    '                    height = reader1Data.GetValue(15)
    '                End If
    '                If reader1Data.IsDBNull(16) = False Then
    '                    spiffAmount = reader1Data.GetValue(16)
    '                End If
    '                If reader1Data.IsDBNull(17) = False Then
    '                    tags = reader1Data.GetString(17)
    '                End If

    '                SetProdInfoForOrderDetail(recID, year, denomination, mintMark, grade, desc, expiresOn, cost, grdSvc, source, coinTypeID, metal, weight, shipWeight, length, width, height, spiffAmount, tags)
    '            End While
    '        End If
    '        reader1Data.Close()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub

    'Public Sub SetProdInfoForOrderDetail(recID As Integer, year As String, denomination As String, mintMark As String, grade As String, desc As String, expiresOn As Date, cost As Decimal, grdSvc As String, source As String, coinTypeID As Integer, metal As String, weight As Decimal, shipWeight As Decimal, length As Decimal, width As Decimal, height As Decimal, spiffAmount As Decimal, tags As String)
    '    Dim sqlData As String
    '    Dim myCommand As New MySqlCommand
    '    Dim connData As New MySqlConnection(GlobalVars.csData)

    '    Try
    '        connData.Open()
    '        myCommand.Connection = connData
    '        sqlData = "UPDATE tbl_OrderDetail SET " &
    '                "Year = @Year, " &
    '                "Denomination = @Denomination, " &
    '                "MintMark = @MintMark, " &
    '                "Grade = @Grade, " &
    '                "Prod_Desc = @Prod_Desc, " &
    '                "ExpiresOn = @ExpiresOn, " &
    '                "Cost = @Cost, " &
    '                "GrdSvc = @GrdSvc, " &
    '                "Source = @Source, " &
    '                "CoinTypeID = @CoinTypeID, " &
    '                "Metal = @Metal, " &
    '                "Weight = @Weight, " &
    '                "ShipWeight = @ShipWeight, " &
    '                "Length = @Length, " &
    '                "Width = @Width, " &
    '                "Height = @Height, " &
    '                "SpiffAmount = @SpiffAmount, " &
    '                "Tags = @Tags " &
    '                "WHERE RecID = @RecID; "
    '        With myCommand
    '            .CommandType = CommandType.Text
    '            .CommandText = sqlData
    '            .Parameters.AddWithValue("@Year", year)
    '            .Parameters.AddWithValue("@Denomination", denomination)
    '            .Parameters.AddWithValue("@MintMark", mintMark)
    '            .Parameters.AddWithValue("@Grade", grade)
    '            .Parameters.AddWithValue("@Prod_Desc", desc)
    '            .Parameters.AddWithValue("@ExpiresOn", expiresOn)
    '            .Parameters.AddWithValue("@Cost", cost)
    '            .Parameters.AddWithValue("@GrdSvc", grdSvc)
    '            .Parameters.AddWithValue("@Source", source)
    '            .Parameters.AddWithValue("@CoinTypeID", coinTypeID)
    '            .Parameters.AddWithValue("@Metal", metal)
    '            .Parameters.AddWithValue("@Weight", weight)
    '            .Parameters.AddWithValue("@ShipWeight", shipWeight)
    '            .Parameters.AddWithValue("@Length", length)
    '            .Parameters.AddWithValue("@Width", width)
    '            .Parameters.AddWithValue("@Height", height)
    '            .Parameters.AddWithValue("@SpiffAmount", spiffAmount)
    '            .Parameters.AddWithValue("@Tags", tags)
    '            .Parameters.AddWithValue("@RecID", recID)
    '        End With
    '        myCommand.ExecuteNonQuery()
    '    Catch ex As Exception

    '    Finally
    '        connData.Close()
    '    End Try
    'End Sub


#End Region
End Module

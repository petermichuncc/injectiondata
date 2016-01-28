Imports System.Data.SqlClient
Imports MongoDB.Driver
Imports MongoDB.Bson
Imports MongoDB.Bson.Serialization.Attributes




Module InjectionDataCollection

    'Declare Global Variables
    Public PressID As String = ""
    Public Wholething As String = ""
    Public Board01ComPort As String = ""
    Public PagingComplete As Boolean = False
    Public pressNumber As Integer = 0
    Public PageMessage As String = ""
    Public AutoState As Integer = 0
    Public CycleTime As Integer = 0
    Public MessageType As String = ""
    Public CellAsset As String = ""
    Dim Offset As Integer = 0
    Dim PagerMessageToSend As String = ""
    Dim PagerNumber As String = ""
    Dim myDateTime As String = ""
    Dim pageID As New List(Of String)()
    Dim currenttime As DateTime = DateTime.Now
    Dim lasttime As DateTime = DateTime.Now


    'MAIN PROGRAM LOOP

    Sub Main()


        CheckForExistingInstance()      'See if this program is already loaded in memory

        Do

            System.Threading.Thread.Sleep(100) 'Wait for .01 seconds before looking

            ReceiveSerialData()

            If MessageType = "A" Then

                GetCellIDfromDB()           'Get machine asset ID number that matches the ordinate machine number coming from the PLC to add to pager message

                GetPagerIDfromDB()          'Get the pager ID number(s) to send messages to all pagers that match the pager class number

                For Each item As String In pageID.ToList

                    PagerNumber = item
                    PagerMessageToSend = CellAsset & Chr(32) & PageMessage

                    SendPageToDB()                                  'Send the message to the paging database

                Next

            ElseIf MessageType = "D" Then

                SendCycleToDB()

            End If

            currenttime = DateTime.Now

            If (Math.Abs(currenttime.Hour - lasttime.Hour)) > 0 Then

                SendSerialData(currenttime.ToString("yyyy-MM-dd HH:mm:ss") & vbCrLf)
                Console.WriteLine("Current Time sent to Injection PLC: " & currenttime.ToString("yyyy-MM-dd HH:mm:ss"))
                lasttime = currenttime

            End If
        Loop
    End Sub


    Sub CheckForExistingInstance()       'See if this program is already running in memory

        If Process.GetProcessesByName _
       (Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MsgBox("Another Instance of this process is already running")
            System.Threading.Thread.Sleep(100)
            Exit Sub

        End If

    End Sub

    Function ReceiveSerialData() As String      ' Receive strings from a serial port. 

        'Dim com4 As IO.Ports.SerialPort = Nothing
        Dim returnStr As String = ""
        MessageType = ""
        PageMessage = ""
        pressNumber = 0
        AutoState = 0
        CycleTime = 0
        Dim ExtractedDateTime As String = ""
        Dim badcomm As String = ""

        Try
            Using com5 As System.IO.Ports.SerialPort = _
                My.Computer.Ports.OpenSerialPort("COM5")

                returnStr = ""

                returnStr = com5.ReadLine()

                If Left(returnStr, 1) <> "A" And Left(returnStr, 1) <> "D" Then

                    Console.WriteLine("Invalid Data: No Message Type. Data: " & returnStr)

                End If

                com5.Close()

            End Using

        Catch ex As Exception        'All Exceptions

            If ex.Message.ToString <> "The operation has timed out." Then
                Console.WriteLine("Error Reading PLC Serial Port: " & ex.Message)
            End If

        End Try

        If Left(returnStr, 1) = "A" Then

            MessageType = Left(returnStr, 1)
            PressID = (Mid(returnStr, 2, 2))
            PageMessage = Mid(returnStr, 4)

        ElseIf Left(returnStr, 1) = "D" Then

            If returnStr.Length = 26 Then
                MessageType = Left(returnStr, 1)
                PressID = (Mid(returnStr, 2, 2))
                AutoState = Val(Mid(returnStr, 4, 1))
                PressID = PressID.TrimStart(New Char() {"0"c})
                ExtractedDateTime = (Mid(returnStr, 5, 26))
                myDateTime = CDate(ExtractedDateTime).ToString("yyyy-MM-dd HH:mm:ss.ff")

            End If

        End If

        Return Nothing

    End Function


    Sub SendSerialData(ByVal InfoToGo As String)

        ' Send strings to a serial port

        System.Threading.Thread.Sleep(10)           'Sleep for .010 second - Delay to prevent COM port error on multiple writes

        Try
            Using com5 As System.IO.Ports.SerialPort = _
                My.Computer.Ports.OpenSerialPort("COM5")

                com5.WriteLine(InfoToGo)                    'Send string to PLC

                com5.Close()

            End Using

            System.Threading.Thread.Sleep(50)      'Give time to close port

        Catch ex As Exception

            Console.WriteLine("Error: Writing PLC Serial Port: " & ex.Message)

        End Try

    End Sub


    Sub GetCellIDfromDB()    'Get CellID (asset ID number) that is related to CellNum (ordinate cell number from PLC) in the SQL DB InjCellID, and put value in CellAsset
        Dim SQLConn As New SqlConnection() 'The SQL Connection
        Dim SQLCmd As New SqlCommand() 'The SQL Command
        Dim SQLdr As SqlDataReader 'The Local Data Store
        Dim ConnString As String = "server=10.87.38.162;database=Master;uid=SA;pwd=password;"
        Dim SQLStr As String = "SELECT CellID from InjCellID Where CellNum=" & Val(PressID) & ";"

        SQLConn.ConnectionString = ConnString 'Set the Connection String
        CellAsset = ""                          'Clear previous Cell Asset ID info from string

        Try

            SQLConn.Open() 'Open the connection
            SQLCmd.Connection = SQLConn 'Sets the Connection to use with the SQL Command
            SQLCmd.CommandText = SQLStr 'Sets the SQL String
            SQLdr = SQLCmd.ExecuteReader 'Gets Data

            SQLdr.Read()            'Read record data
            CellAsset = SQLdr(0)    'Get Cell Asset ID Number from database

            SQLdr.Close() 'Close the SQLDataReader
            SQLConn.Close() 'Close the connection

        Catch ex As SqlClient.SqlException

            Console.WriteLine("Error reading Cell Asset ID from InjCellID")
            Console.WriteLine(ex.Number & " " & ex.Message)

        End Try

    End Sub

    Sub GetPagerIDfromDB()    'Get PagerID that is related to CellNum in the SQL DB InjCellID, and put value in CellAsset

        Dim SQLConn As New SqlConnection() 'The SQL Connection
        Dim SQLCmd As New SqlCommand() 'The SQL Command
        Dim SQLdr As SqlDataReader 'The Local Data Store
        Dim ConnString As String = "server=10.87.38.162;database=Master;uid=SA;pwd=password;"
        Dim SQLStr As String = "SELECT PagerIDNumber from InjPagerNumbers Where PagingClass=1;"


        SQLConn.ConnectionString = ConnString 'Set the Connection String

        pageID.Clear()

        Try

            SQLConn.Open() 'Open the connection
            SQLCmd.Connection = SQLConn 'Sets the Connection to use with the SQL Command
            SQLCmd.CommandText = SQLStr 'Sets the SQL String
            SQLdr = SQLCmd.ExecuteReader 'Gets Data




            While SQLdr.Read()

                pageID.Add(SQLdr("PagerIDNumber").ToString())

            End While

            SQLdr.Close() 'Close the SQLDataReader
            SQLConn.Close() 'Close the connection

        Catch ex As SqlClient.SqlException

            Console.WriteLine("Error reading Pager Numbers from InjPagerNumbers table")
            Console.WriteLine(ex.Number & " " & ex.Message)

        End Try

        

    End Sub

    Sub SendPageToDB()          'Send page to paging database for PagingSQLScanner program to send message to pagers

        Dim SQLConn As New SqlConnection() 'The SQL Connection
        Dim SQLCmd As New SqlCommand() 'The SQL Command
        Dim ConnString As String = "server=10.87.38.162;database=Master;uid=SA;pwd=password;"
        Dim SQLStr As String = "INSERT INTO page1(PressID, pageID, TMStamp) VALUES('" & PagerMessageToSend & "', '" & PagerNumber & "', getdate())"
        Dim exceptionnumber As Integer

        SQLConn.ConnectionString = ConnString       'Set the Connection String

        Try

            SQLConn.Open()                             'Open the connection
            SQLCmd.Connection = SQLConn             'Sets the Connection to use with the SQL Command
            SQLCmd.CommandText = SQLStr                'Sets the SQL String

            SQLCmd.ExecuteNonQuery()                   'Writes Data

            SQLConn.Close()                            'Close the connection

        Catch ex As SqlClient.SqlException

            Console.WriteLine("Error writing record to page table" & " " & PagerNumber & " " & PagerMessageToSend)
            Console.WriteLine(ex.Number & " " & ex.Message)
            exceptionnumber = ex.Number

        End Try

        If exceptionnumber = 0 Then

            PagingComplete = True

            Console.WriteLine("Page written to DB:" & " " & PagerMessageToSend)

        Else

            PagingComplete = False

        End If

    End Sub

    'Public Sub CreateDemo(ByVal database As MongoDatabase) ' // create data for 5 employees
    'End Sub

    Sub SendCycleToDB()

        Dim db As MongoDatabase
        Dim col1 As MongoCollection
        Dim client As MongoClient
        Dim server As MongoServer
        client = New MongoClient("mongodb://10.87.39.103:27017/mongoBlog")
        server = client.GetServer()
        db = server.GetDatabase("mongoBlog")
        server.Connect()
        col1 = db.GetCollection("PressCycles")

        Dim PCycles As MongoCollection(Of BsonDocument) = db.GetCollection("PressCycles")

        AutoState = AutoState.ToString()

        Dim Cycle As BsonDocument = New BsonDocument() _
                                        .Add("PressNumber", PressID) _
                                        .Add("AutoStatus", AutoState.ToString()) _
                                        .Add("CycleTimeStamp", myDateTime)

        PCycles.Insert(Cycle)

        Console.WriteLine("Press cycle written to DB:" & " " & "PressID=" & PressID & " " & myDateTime)

    End Sub

End Module

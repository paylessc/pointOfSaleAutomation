Public Class Form1
    Dim sessionManager As QBPOSFC3Lib.QBPOSSessionManager
    Dim cAppID As String = "noAppID"
    Dim cAppName As String = "QBPOS Custom Reports"
    Dim computerName As String = "server02"
    Dim companyFileName As String = ""
    Dim version As String = ""
    Dim connString = ""
    Dim isInSession = False
    Dim departmentItem As Department
    Dim departmentListSize As Integer
    Dim departmentList(50) As Department
    Dim numOfItems As Integer
    Dim countedItems As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        TextBox1.AppendText("Application load successful." & Environment.NewLine)
        TextBox1.AppendText("Creating session manager." & Environment.NewLine)
        sessionManager = New QBPOSFC3Lib.QBPOSSessionManager()
        TextBox1.AppendText("Session manager created." & Environment.NewLine)
        TextBox1.AppendText("Searching for server..." & Environment.NewLine)
        getServers(sessionManager)
        TextBox1.AppendText("Connecting to database..." & Environment.NewLine)
        connString = "Computer Name=" & computerName & ";Company Data=" & companyFileName & ";Version=" & version '& ";Practice=Yes"
        connectToDatabase()
        TextBox1.AppendText("Requesting Data..." & Environment.NewLine)
        dataRequest()
        SaveToFile()
        Close() 'Before app closes the session is ended then the connection is closed.

    End Sub

    Private Sub getServers(sessionManager As QBPOSFC3Lib.QBPOSSessionManager)
        Dim posserversXML As String

        posserversXML = sessionManager.POSServers(False)
        Dim xmlDoc As MSXML2.DOMDocument40
        Dim objNodeList As MSXML2.IXMLDOMNodeList
        Dim objChild As MSXML2.IXMLDOMNode
        Dim i As Integer
        Dim ret As Boolean
        Dim errorMsg As String

        errorMsg = ""
        xmlDoc = New MSXML2.DOMDocument40

        ret = xmlDoc.loadXML(posserversXML)
        If Not ret Then
            errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
            GoTo ErrHandler
        End If

        Dim server As String
        objNodeList = xmlDoc.getElementsByTagName("POSServer")
        For i = 0 To (objNodeList.length - 1)
            For Each objChild In objNodeList.item(i).childNodes
                If objChild.nodeName = "ServerName" Then
                    server = objChild.text
                End If
                If objChild.nodeName = "CompanyName" Then
                    server = server + " - " + objChild.text
                    companyFileName = objChild.text
                End If
                If objChild.nodeName = "Version" Then
                    server = server + " - " + objChild.text
                    version = objChild.text
                End If
            Next
            TextBox1.AppendText("Server found!" & Environment.NewLine)
        Next

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, vbExclamation, "Error")
        Exit Sub

    End Sub

    Public Sub connectToDatabase()
        sessionManager.OpenConnection(cAppID, cAppName)
        sessionManager.BeginSession(connString)
        TextBox1.AppendText("Database connection successful!" & Environment.NewLine)

    End Sub

    Private Sub dataRequest()
        Dim requestMsgSet As QBPOSFC3Lib.IMsgSetRequest
        Dim responseMsgSet As QBPOSFC3Lib.IMsgSetResponse
        Dim xmlMajorVersion As String = "1"
        Dim xmlMinorVersion As String = "0"

        requestMsgSet = sessionManager.CreateMsgSetRequest(xmlMajorVersion, xmlMinorVersion)
        If (requestMsgSet Is Nothing) Then
            MsgBox("Invalid request to message set object.")
            Close()
            Exit Sub
        End If

        Dim departmentQuery As QBPOSFC3Lib.IDepartmentQuery
        departmentQuery = requestMsgSet.AppendDepartmentQueryRq

        responseMsgSet = sessionManager.DoRequests(requestMsgSet)
        Dim response As QBPOSFC3Lib.IResponse
        Dim statusCode, statusMessage, statusSeverity

        response = responseMsgSet.ResponseList.GetAt(0)
        statusCode = response.StatusCode
        statusMessage = response.StatusMessage
        statusSeverity = response.StatusSeverity

        If (Not statusCode = 0) Then
            MsgBox("Error: " & statusMessage)
            Close()
            Exit Sub
        End If

        If (Not response.Detail Is Nothing) Then
            Dim responseType As Integer
            responseType = response.Type.GetValue
            Dim x As Integer
            Dim departmentRetList As QBPOSFC3Lib.IDepartmentRetList
            departmentRetList = response.Detail
            Dim departmentRet As QBPOSFC3Lib.IDepartmentRet
            departmentListSize = departmentRetList.Count
            For x = 0 To departmentRetList.Count - 1
                departmentRet = departmentRetList.GetAt(x)
                If Not departmentRet.DepartmentCode Is Nothing Then
                    departmentItem = New Department()
                    departmentItem.departmentName = departmentRet.DepartmentName.GetValue
                    departmentItem.departmentCode = departmentRet.DepartmentCode.GetValue
                    departmentItem.departmentID = departmentRet.ListID.GetValue
                    departmentList(x) = departmentItem
                Else
                    departmentItem = New Department()
                    departmentItem.departmentName = departmentRet.DepartmentName.GetValue
                    departmentItem.departmentCode = departmentRet.DepartmentName.GetValue
                    departmentItem.departmentID = departmentRet.ListID.GetValue
                    departmentList(x) = departmentItem
                End If
            Next
        End If

        TextBox1.AppendText("Compiling Data..." & Environment.NewLine)

        Dim itemQuery As QBPOSFC3Lib.IItemInventoryQuery
        requestMsgSet.ClearRequests()
        itemQuery = requestMsgSet.AppendItemInventoryQueryRq

        responseMsgSet = sessionManager.DoRequests(requestMsgSet)

        response = responseMsgSet.ResponseList.GetAt(0)
        statusCode = response.StatusCode
        statusMessage = response.StatusMessage
        statusSeverity = response.StatusSeverity

        If (Not statusCode = 0) Then
            MsgBox("Error: " & statusMessage)
            Close()
            Exit Sub
        End If

        If (Not response.Detail Is Nothing) Then
            Dim responseType As Integer
            responseType = response.Type.GetValue
            Dim itemQuantity As Integer
            Dim x As Integer
            Dim inventoryItemRetList As QBPOSFC3Lib.IItemInventoryRetList
            inventoryItemRetList = response.Detail
            Dim inventoryItemRet As QBPOSFC3Lib.IItemInventoryRet
            numOfItems = inventoryItemRetList.Count
            For x = 0 To inventoryItemRetList.Count - 1
                inventoryItemRet = inventoryItemRetList.GetAt(x)
                For z = 0 To departmentListSize - 1
                    If (Not departmentList(z) Is Nothing) Then
                        If departmentList(z).departmentID = inventoryItemRet.DepartmentListID.GetValue Then
                            itemQuantity = inventoryItemRet.QuantityOnHand.GetValue 'inventoryItemRet.OnHandStore01.GetValue + inventoryItemRet.OnHandStore02.GetValue + inventoryItemRet.OnHandStore03.GetValue + inventoryItemRet.OnHandStore04.GetValue
                            departmentList(z).departmentExtCost += (itemQuantity * inventoryItemRet.Cost.GetValue)
                            countedItems += 1
                            itemQuantity = 0
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        TextBox1.AppendText("Compilation Completed!" & Environment.NewLine)
    End Sub

    Private Sub SaveToFile()
        TextBox1.AppendText("Saving..." & Environment.NewLine)

        Try
            Dim file As New FileProperties
            Using locationFile As New System.IO.StreamReader("C:\myFolder\CustomReportProperties.txt")
                Dim properties As String
                Dim variable As String
                Dim value As String
                Dim x As Integer

                While Not locationFile.EndOfStream

                    properties = locationFile.ReadLine()
                    variable = ""
                    value = ""

                    If properties.Count > 0 Then

                        For x = 0 To properties.Count - 1
                            If Not properties(x) = "=" Then
                                variable += properties(x)

                            ElseIf properties(x) = "=" Then
                                If variable = "Report File Name" Then
                                    value = properties.Substring(x + 1)
                                    While value.StartsWith(" ")
                                        value = value.Remove(0, 1)
                                    End While
                                    While value.EndsWith(" ")
                                        value = value.Remove(value.Count - 1, 1)
                                    End While
                                    file.ReportName = value
                                    Exit For

                                ElseIf variable = "Save Location" Then
                                    value = properties.Substring(x + 1)
                                    While value.StartsWith(" ")
                                        value = value.Remove(0, 1)
                                    End While
                                    While value.EndsWith(" ")
                                        value = value.Remove(value.Count - 1, 1)
                                    End While
                                    file.Location = value
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End While

                Using outputFile As New System.IO.StreamWriter(Convert.ToString(file.Location & file.ReportName & " " & System.DateTime.Now.ToString("yyyy-MM-dd") & ".csv"))
                    outputFile.WriteLine("Department, Ext Cost")
                    For Each department As Department In departmentList
                        If Not department Is Nothing Then
                            outputFile.WriteLine(department.departmentName & "," & department.departmentExtCost)
                        End If
                    Next
                End Using
            End Using
            If My.Computer.FileSystem.GetFiles(file.Location).Count > 7 Then
                Dim filesInDirectory(My.Computer.FileSystem.GetFiles(file.Location).Count) As String
                filesInDirectory = My.Computer.FileSystem.GetFiles(file.Location).ToArray
                Array.Sort(filesInDirectory)
                My.Computer.FileSystem.DeleteFile(filesInDirectory(0))
            End If
        Catch e As Exception
            MsgBox("Could not read save location file: " & e.Message)
        End Try

        TextBox1.AppendText("Done!" & Environment.NewLine)
    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        sessionManager.EndSession()
        sessionManager.CloseConnection()
    End Sub
End Class
Public Class FileProperties
    Dim fileName As String
    Dim saveLocation As String

    Public Sub New()
        Me.fileName = ""
        Me.saveLocation = ""
    End Sub

    Public Property ReportName() As String
        Get
            Return Me.fileName
        End Get
        Set(value As String)
            Me.fileName = value
        End Set
    End Property

    Public Property Location() As String
        Get
            Return Me.saveLocation
        End Get
        Set(value As String)
            Me.saveLocation = value
        End Set
    End Property

End Class

Public Class Department
    Dim name As String
    Dim code As String
    Dim ID As String
    Dim extCost As Double

    Public Sub New()
        Me.name = ""
        Me.extCost = 0.0
        Me.code = ""
        Me.ID = ""
    End Sub

    Public Property departmentName() As String
        Get
            Return Me.name
        End Get
        Set(value As String)
            Me.name = value
        End Set
    End Property

    Public Property departmentID() As String
        Get
            If Not ID Is Nothing Then
                Return Me.ID
            Else
                Return "N/A"
            End If

        End Get
        Set(value As String)
            Me.ID = value
        End Set
    End Property

    Public Property departmentCode() As String
        Get
            If Not code Is Nothing Then
                Return Me.code
            Else
                Return "N/A"
            End If

        End Get
        Set(value As String)
            Me.code = value
        End Set
    End Property

    Public Property departmentExtCost() As Double
        Get
            Return Me.extCost
        End Get
        Set(value As Double)
            Me.extCost = value
        End Set
    End Property

End Class
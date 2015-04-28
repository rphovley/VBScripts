'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''DECLARE VARIABLES''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
''                Identifiers                  ''
'''''''''''''''''''''''''''''''''''''''''''''''''
private pEmail, pName, pID As String
'''''''''''''''''''''''''''''''''''''''''''''''''
''            Payment Information              ''
'''''''''''''''''''''''''''''''''''''''''''''''''
private pInstallPool As Currency
private pFirstJobDate As Date
private repType As String
'''''''''''''''''''''''''''''''''''''''''''''''''
''                Bonus Info                   ''
'''''''''''''''''''''''''''''''''''''''''''''''''
private pNetPromoterQuartile As Integer
private pMaritalStatus As String
private pBenchmarkArray As Array

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''GET/SET VARIABLES''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
''                Identifiers                  ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(value As String)
    pEmail = value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get ID() As Integer
    ID = pID
End Property

Public Property Let ID(value As Integer)
    pID = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''            Payment Information              ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get InstallPool() As Currency
    InstallPool = pInstallPool
End Property

Public Property Let InstallPool(value As Currency)
    pInstallPool = value
End Property

Public Property Get FirstJobDate() As Date
    FirstJobDate = pFirstJobDate
End Property

Public Property Let FirstJobDate(value As Date)
    pFirstJobDate = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                Bonus Info                   ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get NetPromoterQuartile() As String
    NetPromoterQuartile = pNetPromoterQuartile
End Property

Public Property Let NetPromoterQuartile(value As String)
    pNetPromoterQuartile = value
End Property

Public Sub setNetPromoter(ByRef weatherData As Dictionary)
    Dim isArray, statusArray, stateArray As Variant
    Dim weatherRep As cWeatherData
    isArray = Array("Site Survey Complete", "Design Complete", "Application Complete", _
        "Submitted", "Rejected", "Received")
    statusArray = Array("Sales", "Permit")
    stateArray  = Array("NY", "NJ", "DE", "MD", "CT", "MA")
    Dim isInclementState As Boolean

    On Error Resume Next
    Set weatherRep = weatherData.Item(Me.RepEmail)
    If NOT weatherRep is nothing Then
        isInclementState = True
    Else
        For Each sState In stateArray
            If Me.States = sState Then
                isInclementState = True
            End IF
        Next
    End If
        'Loops through backend statuses that trigger backend'
            If Me.Status = "Permit" Then
                For Each arrayStatus In isArray
                    'if it is a correct backend status, return true'
                    If NOT isInclementState AND arrayStatus = Me.SubStatus Then
                        Me.IsSurveyComplete = True
                        Exit For
                    ElseIf isInclementState AND arrayStatus = Me.SubStatus OR "Site Survey Scheduled" = Me.Substatus Then
                        Me.IsSurveyComplete = True
                        Exit For
                    End If

                Next arrayStatus

            ElseIf Me.Status = "Sales" Then
                Me.IsSurveyComplete = False
            ElseIf NOT Me.IsCancelled Then
                Me.IsSurveyComplete = True
            End If
End Sub

Public Property Get MaritalStatus() As String
    MaritalStatus = pMaritalStatus
End Property

Public Property Let MaritalStatus(value As String)
    pMaritalStatus = value
End Property

Public Property Get BenchmarkArray() As String
    BenchmarkArray = pBenchmarkArray
End Property

Public Property Let BenchmarkArray(value As String)
    pBenchmarkArray = value
End Property

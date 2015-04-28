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
Public Property Get NetPromoterQuartile() As Integer
    NetPromoterQuartile = pNetPromoterQuartile
End Property

Public Property Let NetPromoterQuartile(value As Integer)
    pNetPromoterQuartile = value
End Property

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
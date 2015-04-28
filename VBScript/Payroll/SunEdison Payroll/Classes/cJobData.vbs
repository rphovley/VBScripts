'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''DECLARE VARIABLES''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
''                Identifiers                  ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Private pJobID, pRepEmail, pCustomer As String

'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Statuses                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Set pJobStatusDict = New Scripting.Dictionary
    
'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Payments                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Set pPaymentDict   = New Scripting.Dictionary

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Dates                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Private pCreatedDate As Date
Private pDaysSinceCreated As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Other                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Private pkW As Double
Private pTotalPay As Currency



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''GET/SET VARIABLES''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
''                Identifiers                  ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get JobID() As String
    JobID = pJobID
End Property

Public Property Let JobID(value As String)
    pJobID = value
End Property

Public Property Get RepEmail() As String
    RepEmail = pRepEmail
End Property

Public Property Let RepEmail(value As String)
    pRepEmail = value
End Property

Public Property Get Customer() As String
    Customer = pCustomer
End Property

Public Property Let Customer(value As String)
    pCustomer = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Statuses                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get jobStatusDict() As Dictionary
    jobStatusDict = pJobStatusDict
End Property

Public Property Let jobStatusDict(value As Dictionary)
    pJobStatusDict = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Payments                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get PaymentDict() As Dictionary
    PaymentDict = pPaymentDict
End Property

Public Property Let PaymentDict(value As Dictionary)
    pPaymentDict = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Dates                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get CreatedDate() As Date
    CreatedDate = pCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    pCreatedDate = value
End Property

Public Property Get DaysSinceCreated() As Integer
    DaysSinceCreated = pDaysSinceCreated
End Property

Public Property Let DaysSinceCreated(value As Integer)
    pDaysSinceCreated = value
End Property

Public Sub setDaysSinceCreated()
	pDaysSinceCreated = DateDiff("d",pCreatedDate, Now())
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Other                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get kW() As Double
    kW = pkW
End Property

Public Property Let kW(value As Double)
    pkW = value
End Property

Public Property Get totalPay() As Currency
    totalPay = pTotalPay
End Property

Public Property Let totalPay(value As Currency)
    pTotalPay = value
End Property



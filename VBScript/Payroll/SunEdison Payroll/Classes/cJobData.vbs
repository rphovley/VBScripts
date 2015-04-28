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
'Unique job ID'
Public Property Get JobID() As String
    JobID = pJobID
End Property

Public Property Let JobID(value As String)
    pJobID = value
End Property

'Each rep email is unique'
Public Property Get RepEmail() As String
    RepEmail = pRepEmail
End Property

Public Property Let RepEmail(value As String)
    pRepEmail = value
End Property

'This is the customer's name'
Public Property Get Customer() As String
    Customer = pCustomer
End Property

Public Property Let Customer(value As String)
    pCustomer = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Statuses                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''
'A dictionary containing the job status as the key, and a boolean as the value'
Public Property Get jobStatusDict() As Dictionary
    jobStatusDict = pJobStatusDict
End Property

Public Property Let jobStatusDict(value As Dictionary)
    pJobStatusDict = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                 Payments                    ''
'''''''''''''''''''''''''''''''''''''''''''''''''

' A dictionary containing the date of the payment          '         
' as the key, and the amount of the payment as the value   '
Public Property Get paymentDict() As Dictionary
    paymentDict = pPaymentDict
End Property

Public Property Let paymentDict(value As Dictionary)
    pPaymentDict = value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Dates                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
'The date the job is created'
Public Property Get CreatedDate() As Date
    CreatedDate = pCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    pCreatedDate = value
End Property

'The difference between today's date and the created date'
Public Property Get DaysSinceCreated() As Integer
    DaysSinceCreated = pDaysSinceCreated
End Property

Public Property Let DaysSinceCreated(value As Integer)
    pDaysSinceCreated = value
End Property

'Method to set the days since created value'
Public Sub setDaysSinceCreated()
	pDaysSinceCreated = DateDiff("d",pCreatedDate, Now())
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
''                  Other                      ''
'''''''''''''''''''''''''''''''''''''''''''''''''
'The Kilowatts for the job'
Public Property Get kW() As Double
    kW = pkW
End Property

Public Property Let kW(value As Double)
    pkW = value
End Property

'The sum of values of the key/value pairs in the PaymentDict'
Public Property Get totalPay() As Currency
    totalPay = pTotalPay
End Property

Public Property Let totalPay(value As Currency)
    pTotalPay = value
End Property



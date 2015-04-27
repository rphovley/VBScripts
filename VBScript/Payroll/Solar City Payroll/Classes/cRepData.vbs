private pName, pEmail As String
private pID, pPayScaleID As Integer
private pKwSum As Double
private pIsBlackList, pIsInactive, pIsSlider, pIsMarketing As Boolean
private pMarketingRate, pInstallPool As Currency
private pFirstJobDate, pStartSliderDate, pMarkStartDate, pMarkEndDate As Date
private pSalesThisWeek as Integer

' Get/Set methods'
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

Public Sub setIsMarketing()
    If Now() > pMarkStartDate AND Now() < pMarkEndDate Then
        pIsMarketing = True
    End If
End Sub

Public Property Let IsMarketing(value As Boolean)
    pIsMarketing = value
End Property

Public Property Get IsMarketing() As Boolean
    IsMarketing = pIsMarketing
End Property

Public Property Get MarketingRate() As Currency
    MarketingRate = pMarketingRate 
End Property

Public Property Let MarketingRate(value As Currency)
    pMarketingRate = value
End Property

Public Property Get MarkEndDate() As Date
    MarkEndDate = pMarkEndDate
End Property

Public Property Let MarkEndDate(value As Date)
    pMarkEndDate = value
End Property

Public Property Get MarkStartDate() As Date
    MarkStartDate = pMarkStartDate
End Property

Public Property Let MarkStartDate(value As Date)
    pMarkStartDate = value
End Property

Public Property Let SalesThisWeek(value as Integer)
	pSalesThisWeek = value
End Property

Public Property Get SalesThisWeek() As Integer
	SalesThisWeek = pSalesThisWeek
End Property

Public Property Let IsSlider(value As Boolean)
    pIsSlider = value
End Property

Public Property Get IsSlider() As Boolean
    IsSlider = pIsSlider
End Property

Public Property Get StartSliderDate() As Date
    StartSliderDate = pStartSliderDate
End Property

Public Property Let StartSliderDate(value As Date)
    pStartSliderDate = value
End Property

Public Property Let KwSum(value As Double)
    pKwSum = value
End Property

Public Property Get KwSum() As Double
    KwSum = pKwSum
End Property

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

Public Property Get PayScaleID() As Integer
    PayScaleID = pPayScaleID
End Property

Public Property Let PayScaleID(value As Integer)
    pPayScaleID = value
End Property

Public Property Let IsBlackList(value As Boolean)
    pIsBlackList = value
End Property

Public Property Get IsBlackList() As Boolean
    IsBlackList = pIsBlackList
End Property

Sub setIsBlackList(ByVal val As String)
    If val = "Y" Then
        pIsBlackList = True
    End If
End Sub

Public Property Let IsInactive(value As Boolean)
    pIsInactive = value
End Property

Public Property Get IsInactive() As Boolean
    IsInactive = pIsInactive
End Property

Sub setIsInactive(ByVal val As String)
    If val = "Y" Then
        pIsInactive = True
    End If
End Sub


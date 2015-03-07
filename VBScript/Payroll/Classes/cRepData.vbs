private pName, pEmail As String
private pID, pPayScaleID As Integer
private pKwSum As Double
private pIsBlackList, pIsInactive, pIsNewRep, pIsSlider As Boolean
private pStartSliderDate As Date

' Get/Set methods'
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

Public Property Let IsNewRep(value As Boolean)
    pIsNewRep = value
End Property

Public Property Get IsNewRep() As Boolean
    IsNewRep = pIsNewRep
End Property


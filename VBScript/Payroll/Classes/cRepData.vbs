private pName, pEmail As String
private pID, pPayScaleID As Integer
private pIsBlackList, pIsInactive As Boolean

' Get/Set methods'
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

Public Property Get RPayScaleID() As Integer
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
Private pID As Integer
Private pName As String
Private pRate As Variant

Public Property Get ID() As Integer
    ID = pID
End Property

Public Property Let ID(value As Integer)
    pID = value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Rate() As Variant
    Rate = pRate
End Property

Public Property Let Rate(value As Variant)
    pRate = value
End Property

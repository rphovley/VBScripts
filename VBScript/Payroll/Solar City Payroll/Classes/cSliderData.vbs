Private pID As Integer
Private pFirstQuartileMax, pSecondQuartileMax, pThirdQuartileMax, pFourthQuartileMax As Integer
Private pFirstQuartileKW, pSecondQuartileKW, pThirdQuartileKW, pFourthQuartileKW As Currency
Public Property Get ID() As Integer
    ID = pID
End Property

Public Property Let ID(value As Integer)
    pID = value
End Property

Public Property Get FirstQuartileMax() As Integer
    FirstQuartileMax = pFirstQuartileMax
End Property

Public Property Let FirstQuartileMax(value As Integer)
    pFirstQuartileMax = value
End Property

Public Property Get SecondQuartileMax() As Integer
    SecondQuartileMax = pSecondQuartileMax
End Property

Public Property Let SecondQuartileMax(value As Integer)
    pSecondQuartileMax = value
End Property

Public Property Get ThirdQuartileMax() As Integer
    ThirdQuartileMax = pThirdQuartileMax
End Property

Public Property Let ThirdQuartileMax(value As Integer)
    pThirdQuartileMax = value
End Property

Public Property Get FourthQuartileMax() As Integer
    FirstQuartileMax = pFirstQuartileMax
End Property

Public Property Let FourthQuartileMax(value As Integer)
    pFourthQuartileMax = value
End Property

Public Property Get FirstQuartileKW() As Currency
    FirstQuartileKW = pFirstQuartileKW
End Property

Public Property Let FirstQuartileKW(value As Currency)
    pFirstQuartileKW = value
End Property

Public Property Get SecondQuartileKW() As Currency
    SecondQuartileKW = pSecondQuartileKW
End Property

Public Property Let SecondQuartileKW(value As Currency)
    pSecondQuartileKW = value
End Property

Public Property Get ThirdQuartileKW() As Currency
    ThirdQuartileKW = pThirdQuartileKW
End Property

Public Property Let ThirdQuartileKW(value As Currency)
    pThirdQuartileKW = value
End Property

Public Property Get FourthQuartileKW() As Currency
    FourthQuartileKW = pFourthQuartileKW
End Property

Public Property Let FourthQuartileKW(value As Currency)
    pFourthQuartileKW = value
End Property



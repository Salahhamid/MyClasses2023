VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pkey As Long
Private pName As String
Private pChildren As Collection
Public Property Get key() As Long
    key = pkey
End Property
Public Property Get Name() As String
    Name = pName
End Property
Public Property Get children() As Collection
    Set children = pChildren
End Property
Public Property Let key(p As Long)
    pkey = p
End Property
Public Function init(k As Long, sName As String) As cMyClass
    pkey = k
    pName = sName
    Set pChildren = New Collection
    Set init = Me
End Function

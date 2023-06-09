VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVBAProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' v2.01
' one of these for each project in the workbook
Private pWorkBook As Workbook
Private pModules As Collection

Public Property Get Name() As String
    Name = pWorkBook.VBProject.Name
End Property
Public Property Get wBook() As Workbook
    Set wBook = pWorkBook
End Property

Public Property Get modules() As Collection
    Set modules = pModules
End Property
Public Sub init(wb As Workbook)
    Set pWorkBook = wb

End Sub
Public Sub tearDown()
    Dim vm As cVBAmodule
    For Each vm In pModules
        vm.tearDown
    Next vm
    Set pModules = Nothing
End Sub

Private Sub Class_Initialize()
    Set pModules = New Collection
End Sub

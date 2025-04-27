VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
    ID As Long
    Description As String * 50
    Number As String * 10
    DateOfError As Date
    DateTime As Date
    FormName As String * 20
    ReportName As String * 20
    UserName As String * 20
End Type

Private udtProps As DisplayProps

Public Property Get ID() As Long
  ID = udtProps.ID
End Property

Friend Property Let ID(Value As Long)
  udtProps.ID = Value
End Property

Friend Property Let Description(Value As String)
  udtProps.Description = Value
End Property
Public Property Get Description() As String
  Description = Trim$(udtProps.Description)
End Property

Friend Property Let Number(Value As String)
  udtProps.Number = Value
End Property
Public Property Get Number() As String
  Number = Trim$(udtProps.Number)
End Property

Friend Property Let DateOfError(Value As String)
  udtProps.DateOfError = Value
End Property
Public Property Get DateOfError() As String
  DateOfError = Trim$(udtProps.DateOfError)
End Property
Friend Property Let DateTime(Value As String)
  udtProps.DateTime = Value
End Property
Public Property Get DateTime() As String
  DateTime = Trim$(udtProps.DateTime)
End Property

Friend Property Let FormName(Value As String)
  udtProps.FormName = Value
End Property
Public Property Get FormName() As String
  FormName = Trim$(udtProps.FormName)
End Property
Friend Property Let ReportName(Value As String)
  udtProps.ReportName = Value
End Property
Public Property Get ReportName() As String
  ReportName = Trim$(udtProps.ReportName)
End Property
Friend Property Let UserName(Value As String)
  udtProps.UserName = Value
End Property
Public Property Get UserName() As String
  UserName = Trim$(udtProps.UserName)
End Property



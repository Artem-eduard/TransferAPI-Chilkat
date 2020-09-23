VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Req As WinHttp.WinHttpRequest
Public Function SendPost(ByRef the_sURL As String, ByVal bodyStr As String) As String

Set Req = New WinHttp.WinHttpRequest
    With Req
        .Open "POST", the_sURL, Async:=False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Accept", "*/*"
        
        .Send """iotBilling"""
        
        'Note: Normally you don't include all of this whitespace, but
        'we'll use it in this example:
        
      '  Label1.Caption = CStr(.Status) & " " & .StatusText & vbNewLine _
                       & .GetAllResponseHeaders() & vbNewLine _
                       & String$(40, "-") & vbNewLine _
                       & .ResponseText
       SendPost = .ResponseText
    End With
    
    
    
    
End Function


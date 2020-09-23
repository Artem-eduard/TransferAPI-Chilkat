VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function Authenticate(ByRef the_sURL As String, ByVal bodyStr As String) As String

Dim rest As New ChilkatRest

' Connect to the REST server.
Dim bTls As Long
bTls = 1
Dim port As Long
port = 443
Dim bAutoReconnect As Long
bAutoReconnect = 1
Dim success As Long
success = rest.Connect(the_sURL, port, bTls, bAutoReconnect)
If (success <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Function
End If
success = rest.AddHeader("Content-Type", "application/json; charset=UTF-8")

success = rest.AddQueryParam("api-version", "1.1")

Dim responseJson As String

responseJson = rest.FullRequestString("POST", "/Authentication/Contact/Authenticate/Site:demo3/ContactCode:40000022", bodyStr)

If (rest.LastMethodSuccess <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Function
End If

Debug.Print responseJson
SendPost = responseJson

    
End Function


Public Function Post(ByRef the_sURL As String, ByVal tokenStr As String, ByVal URIStr As String, ByVal URIvalueStr As String, ByVal bodyStr As String) As String

Dim rest As New ChilkatRest

' Connect to the REST server.
Dim bTls As Long
bTls = 1
Dim port As Long
port = 443
Dim bAutoReconnect As Long
bAutoReconnect = 1
Dim success As Long
success = rest.Connect(the_sURL, port, bTls, bAutoReconnect)
If (success <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Function
End If


success = rest.AddHeader("Content-Type", "application/json; charset=UTF-8")
success = rest.AddQueryParam("api-version", "1.1")
success = rest.AddHeader("Authorization", tokenStr)
Dim responseJson As String
Dim sbResponseBody As New ChilkatStringBuilder

responseJson = rest.FullRequestSb("POST", URIStr & ":" & URIvalueStr, bodyStr, sbResponseBody)

If (rest.LastMethodSuccess <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Function
End If

Debug.Print responseJson
SendPost = responseJson

    
End Function

Public Function JsonToRecordset(ByRef JSON As String) As ADODB.Recordset
    Dim jbRoot As JsonBag
    Dim RS As New ADODB.Recordset
    Dim FieldNames As Variant
    Dim ColIndex As Long
    Dim jbField As JsonBag
    Dim DataType As ADODB.DataTypeEnum
    Dim Values As Variant
    Dim RowIndex As Long
    Dim jbRow As JsonBag
    Dim jbCol As JsonBag

    Set JsonToRecordset = New ADODB.Recordset
    With JsonToRecordset
        .CursorLocation = adUseClient
        Set jbRoot = New JsonBag
        jbRoot.JSON = JSON
        With .Fields
            ReDim FieldNames(1 To jbRoot("Fields").Count)
            For ColIndex = 1 To jbRoot("Fields").Count
                Set jbField = jbRoot("Fields")(ColIndex)
                FieldNames(ColIndex) = jbField("Name")
                Select Case jbField("Type")
                    'No support for Single and Double types so we punt and just
                    'make them all Double:
                    Case "REAL"
                        DataType = adDouble
                    
                    'Hazard:
                    '
                    '   The crude "INTEGER" type can be anything from a Byte to a
                    '   LongLong (64-bit integer).
                    '
                    '   To cope here we'll stuff them all into adBigInt (LongLong).
                    '   Since VB6 does not directly support LongLong you may run
                    '   into issues you wil have to work around if values exceed
                    '   the range of a Long.
                    '
                    'No support for strongly typed integer sizes.  Just punt and
                    'make them all LongLong:
                    Case "INTEGER"
                        DataType = adBigInt

                    'Just punt and use the adLongVarWChar data type with the
                    'maximum length to handle the crude attempt at a "TEXT" data
                    'type:
                    Case "TEXT"
                        DataType = adLongVarWChar

                    'Probably won't happen?
                    '
                    '   JSON doesn't really have a decent way to represent a BLOB,
                    '   so if you have any BLOB data you are sort of screwed anyway.
                    '
                    'Just punt and use the adLongVarBinary data type with the
                    'maximum length to handle its crude attempt at a "BLOB" data
                    'type which you'll probably NEVER get anyway:
                    Case "BLOB"
                        DataType = adLongVarBinary

                    Case Else
                        Err.Raise 5 'Just blow up.  The crude DBMS doesn't seem to
                                    'have any more anyway.
                End Select
                If DataType = adLongVarWChar Or DataType = adLongVarBinary Then
                    .Append jbField("Name"), _
                            DataType, _
                            &H7FFFFFFF, _
                            IIf(jbField("Nullable"), adFldIsNullable, 0)
                Else
                    .Append jbField("Name"), _
                            DataType, _
                            , _
                            IIf(jbField("Nullable"), adFldIsNullable, 0)
                End If
            Next
        End With
        .Open
        For ColIndex = 1 To jbRoot("Fields").Count
            Set jbField = jbRoot("Fields")(ColIndex)
            If jbField("PrimaryKey") Then
                .Fields(jbField("Name")).Properties("Optimize").Value = True
                Exit For 'PrimaryKey, so there can only be one anyway.
            End If
        Next
        ReDim Values(1 To UBound(FieldNames))
        For RowIndex = 1 To jbRoot("RowsCols").Count
            Set jbRow = jbRoot("RowsCols")(RowIndex)
            For ColIndex = 1 To jbRow.Count
                Values(ColIndex) = jbRow(ColIndex)
            Next
            .AddNew FieldNames, Values
        Next
        .MoveFirst
    End With
End Function


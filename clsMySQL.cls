VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Connection As New ADODB.Connection
Public Command As New ADODB.Command
Public CommandExecuteOptions As Long
Public RecordsAffected

Public Function Connect(Optional sConnectionString As String, Optional sUser As String, Optional sPassword As String, Optional lngOpts As Long = -1)
    Connection.Open sConnectionString, sUser, sPassword, lngOpts
End Function

Public Function Query(ByVal sQueryString As String, ParamArray Params() As Variant) As Recordset
    Dim Arr As Variant
    Dim i As Integer
    Dim j As Integer
    Dim ParamsProvidedCnt As Integer
    
    If UBound(Params) = 0 Then
        If TypeName(Params(0)) = "Variant()" Then
            Arr = Params(0)
        End If
    Else
        Arr = Params
    End If
    
    Set Command.ActiveConnection = Connection
    Command.CommandText = sQueryString

    If Command.Parameters.Count > 0 Then
        If IsEmpty(Arr) Then
            ParamsProvidedCnt = 0
        Else
            ParamsProvidedCnt = UBound(Arr) + 1
        End If
        If ParamsProvidedCnt < Command.Parameters.Count Then
            Err.Raise vbObjectError + 1, "clsMySQL by TQProjekt(R)", _
                Description:="Error! " & Command.Parameters.Count - ParamsProvidedCnt & _
                " paramtere(s) missing in query:" & Chr(10) & Command.CommandText
        End If
    
        j = LBound(Arr)
        For i = 0 To Command.Parameters.Count - 1
            Command.Parameters(i).Value = Arr(j)
            j = j + 1
        Next i
        
    End If
    
    Set Query = Command.Execute(RecordsAffected)
    
End Function








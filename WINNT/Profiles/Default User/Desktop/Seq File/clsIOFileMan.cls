VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIOFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private intFileID As Integer
Private strFilePath As String
Private strFileName As String
Private lngCurrentRow As Long
Private lngRowCount As Long
Public Enum enmWriteType
    AppendData = 0
    InsertData = 1
    ReplaceData = 2
End Enum

Private Sub Class_Initialize()
    intFileID = 0
End Sub

Public Sub OpenFile(strFileLocation As String, pstrFileName As String)
    intFileID = FreeFile

    If Right(strFileLocation, 1) = "\" Then
        strFilePath = strFileLocation
    Else
        strFilePath = strFileLocation & "\"
    End If
    strFileName = pstrFileName
    
    If Len(Dir(strFilePath & strFileName)) > 0 Then SetRowCount

End Sub

Public Sub MoveFirst()
    Close intFileID
    If intFileID > 0 Then Open strFilePath & strFileName For Input As intFileID
    lngCurrentRow = 0
End Sub

Private Sub SetRowCount()
    Dim strTemp As String
    
    Close intFileID
    
    lngRowCount = 0
    Open strFilePath & strFileName For Input As intFileID
    Do While Not EOF(intFileID)
        Line Input #1, strTemp
        lngRowCount = lngRowCount + 1
    Loop
    MoveFirst
End Sub

Public Function RowCount() As Long
    RowCount = lngRowCount
End Function

Public Function CurrentRow() As Long
    CurrentRow = lngCurrentRow
End Function


Public Sub MoveToRow(lngRowID As Long)
    Dim lngTemp As Long
    Dim strTemp As String
    
    If (lngRowID > lngRowCount) Or (intFileID = 0) Then Exit Sub
    
    lngTemp = 0
    
    MoveFirst
    Do While (Not EOF(intFileID)) And (lngTemp < lngRowID)
        Line Input #intFileID, strTemp
        lngTemp = lngTemp + 1
    Loop
    
    lngCurrentRow = lngTemp
    
End Sub

Public Sub MoveNext()
    MoveToRow lngCurrentRow + 1
End Sub

Public Sub MovePrevious()
    MoveToRow lngCurrentRow - 1
End Sub

Public Sub MoveLast()
    MoveToRow lngRowCount
End Sub

Public Function GetDataFromRow(Optional lngRowID As Long = 0) As String
    Dim strTemp As String
    
    If (lngRowID > lngRowCount) Or (intFileID = 0) Then Exit Function
    
    lngCurrentRow = lngRowID + Abs((lngCurrentRow * (lngRowID = 0)))
    
    MoveToRow lngCurrentRow - 1
    
    Line Input #intFileID, strTemp
    
    lngCurrentRow = lngCurrentRow + 1
    
    GetDataFromRow = strTemp
    
End Function

Public Sub DeleteRow(Optional lngRowID As Long = 0)
    If lngRowCount = 0 Then Exit Sub
    WriteData ReplaceData, "", lngRowID
    MoveToRow (Abs((lngCurrentRow - 1) = 0) + (lngCurrentRow - 1))
End Sub

Public Sub WriteData(intWriteType As enmWriteType, strData As String, Optional lngRowID As Long = 0)
    Dim intWriteFileID As Integer
    Dim lngCounter As Long
    Dim lngWriteToRowID As Long
    Dim strTemp As String
    
    If (lngRowID > lngRowCount) Or (intFileID = 0) Then Exit Sub
    
    lngCounter = 0
    
    intWriteFileID = FreeFile
    
    Open strFilePath & "tmp" & strFileName For Output As intWriteFileID
    
    If intWriteType = AppendData Then
        lngWriteToRowID = lngRowCount
    Else
        lngWriteToRowID = lngRowID + Abs((lngCurrentRow * (lngRowID = 0))) - 1
    End If
    
    MoveFirst
    
    Do While lngCounter < lngWriteToRowID
        Line Input #intFileID, strTemp
        Print #intWriteFileID, strTemp
        lngCounter = lngCounter + 1
    Loop
    
    If Len(Trim(strData)) > 0 Then Print #intWriteFileID, strData
    
    If intWriteType = ReplaceData Then Line Input #intFileID, strTemp
    
    Do While Not EOF(intFileID)
        Line Input #intFileID, strTemp
        Print #intWriteFileID, strTemp
    Loop

    Close
    
    Kill strFilePath & strFileName
    Name strFilePath & "tmp" & strFileName As strFilePath & strFileName
    
    OpenFile strFilePath, strFileName
    
    lngCurrentRow = lngWriteToRowID + 1
    
    MoveToRow lngWriteToRowID + 1

End Sub

Private Sub Class_Terminate()
    Close
End Sub

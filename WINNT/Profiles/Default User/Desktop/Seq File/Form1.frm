VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sequential File Manager"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2310
      TabIndex        =   9
      Top             =   2610
      Width           =   1395
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   1620
      Width           =   4965
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Delete Record"
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Append Record"
      Height          =   375
      Index           =   6
      Left            =   3080
      TabIndex        =   7
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Insert Record"
      Height          =   375
      Index           =   5
      Left            =   1600
      TabIndex        =   6
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Replace Record"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Move Last"
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   4
      Top             =   1140
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Move Next"
      Height          =   375
      Index           =   2
      Left            =   3080
      TabIndex        =   3
      Top             =   1140
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Move Previous"
      Height          =   375
      Index           =   1
      Left            =   1600
      TabIndex        =   2
      Top             =   1140
      Width           =   1395
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Move First"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "New data:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1650
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Data:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblData 
      Height          =   225
      Left            =   660
      TabIndex        =   12
      Top             =   720
      Width           =   5265
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current Record:"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   390
      Width           =   2415
   End
   Begin VB.Label lblCount 
      Caption         =   "Record Count:"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objIOFile As clsIOFileMan

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
    With objIOFile
        Select Case Index
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
            Case 2
                .MoveNext
            Case 3
                .MoveLast
            Case 4
                .WriteData ReplaceData, txtData.Text
            Case 5
                .WriteData InsertData, txtData.Text
            Case 6
                .WriteData AppendData, txtData.Text
            Case 7
                .DeleteRow
        End Select
    End With
    
    PopulateFields
    
    txtData.Text = ""
End Sub

Private Sub Form_Load()
    Set objIOFile = New clsIOFileMan
    With objIOFile
        .OpenFile App.Path, "File.txt"
        .MoveFirst
    End With
    PopulateFields
End Sub

Private Sub PopulateFields()
    With objIOFile
        lblCount.Caption = "Record count: " & .RowCount
        lblData.Caption = .GetDataFromRow
        lblCurrent.Caption = "Current record: " & .CurrentRow
    End With
End Sub

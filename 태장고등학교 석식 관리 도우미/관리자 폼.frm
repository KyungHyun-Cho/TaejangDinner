VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  '단일 고정
   Caption         =   "관리자 정보 설정 및 조회"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5370
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5370
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "선택 삭제"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "전체 삭제"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "검색"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2655
      Begin VB.TextBox Text3 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   840
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "검색"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "이름 :"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "관리자 List"
      ClipControls    =   0   'False
      Height          =   2415
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List1 
         Appearance      =   0  '평면
         Height          =   1710
         Left            =   120
         Style           =   1  '확인란
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   180
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "총 관리자 수 :"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "관리자 추가"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2655
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   840
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   840
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "권한 :"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "이름 :"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "바코드 :"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   660
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command12_Click()

End Sub

Private Sub Command2_Click()
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
C = B.readall
B.Close
If List1.ListIndex = "-1" Then MsgBox "리스트를 선택해주세요.", vbCritical, "Error": Exit Sub
If MsgBox("선택하신 리스트가 삭제됩니다. 계속하시겠습니까?", vbInformation + vbYesNo, "확인") = vbYes Then
If MsgBox("삭제되면 다시는 복구할 수 없습니다." & vbCrLf & "정말로 진행하시겠습니까?", vbCritical + vbYesNo, "2차 확인") = vbYes Then
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then
C = Replace(C, "" & List1.List(i) & "", "")
List1.RemoveItem i
End If
DoEvents
Next i
End If
End If
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase3.db", True)
 B.Write (C)
 B.Close
Label6.Caption = List1.ListCount
End Sub

Private Sub Command3_Click()
Dim Ar, B, C



If MsgBox("전체 리스트가 삭제됩니다. 계속하시겠습니까?", vbInformation + vbYesNo, "확인") = vbYes Then
If MsgBox("삭제되면 다시는 복구할 수 없습니다." & vbCrLf & "정말로 진행하시겠습니까?", vbCritical + vbYesNo, "2차 확인") = vbYes Then

 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase3.db", True)
List1.Clear
B.Write (vbCrLf)
B.Close
End If
End If

Label6.Caption = List1.ListCount
End Sub

Private Sub Form_Load()
On Error Resume Next
Combo1.AddItem "Normal Admin"
Combo1.AddItem "Super Admin"
Combo1.ListIndex = 0
On Error Resume Next
Dim Ar, B, C, D
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 3 Then List1.AddItem D: D = ""
DoEvents
Next i
Label6.Caption = List1.ListCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
Form1.UserData = B.readall
 B.Close
Form7.DoingA = False
Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
List1.AddItem "" & Text2.Text & "/" & Text1.Text & "/" & Val(Combo1.ListIndex) + 1 & "/"
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 8, True)
B.WriteLine "" & Text2.Text & "/" & Text1.Text & "/" & Val(Combo1.ListIndex) + 1 & "/"
 B.Close
Label6.Caption = List1.ListCount
Text1.Text = ""
Text2.Text = ""
Combo1.ListIndex = "0"
Text1.SetFocus
End If
End Sub

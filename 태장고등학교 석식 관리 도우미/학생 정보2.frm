VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "석식 관리 프로그램 3.3"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15780
   LinkTopic       =   "Form7"
   ScaleHeight     =   9180
   ScaleWidth      =   15780
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin VB.CheckBox Check1 
      Caption         =   "NetWork Mode"
      Enabled         =   0   'False
      Height          =   180
      Left            =   0
      TabIndex        =   29
      Top             =   720
      Width           =   1935
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2640
      Top             =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   9000
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   13575
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "현재 급식이 닫힌 상태입니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   12840
      End
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   2280
      Top             =   1680
   End
   Begin VB.Timer Timer5 
      Left            =   1920
      Top             =   1680
   End
   Begin VB.Timer Timer4 
      Interval        =   50
      Left            =   1560
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      Height          =   9135
      Left            =   10920
      TabIndex        =   9
      Top             =   120
      Width           =   9015
      Begin VB.Label M0 
         AutoSize        =   -1  'True
         Caption         =   "0 . 관리자 메뉴 종료"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   22
         Top             =   8520
         Width           =   4650
      End
      Begin VB.Label M9 
         AutoSize        =   -1  'True
         Caption         =   "9 . 석식 받기 처리"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   21
         Top             =   7320
         Width           =   4155
      End
      Begin VB.Label M8 
         AutoSize        =   -1  'True
         Caption         =   "8 . 석식 종료"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   20
         Top             =   6720
         Width           =   2985
      End
      Begin VB.Label M7 
         AutoSize        =   -1  'True
         Caption         =   "7 . 석식 시작"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   19
         Top             =   6120
         Width           =   2985
      End
      Begin VB.Label M6 
         AutoSize        =   -1  'True
         Caption         =   "6 . 로그창 열기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   18
         Top             =   5520
         Width           =   3480
      End
      Begin VB.Label M5 
         AutoSize        =   -1  'True
         Caption         =   "5 . 3학년 데이터베이스 다운로드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   7410
      End
      Begin VB.Label M4 
         AutoSize        =   -1  'True
         Caption         =   "4 . 2학년 데이터베이스 다운로드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   16
         Top             =   4320
         Width           =   7410
      End
      Begin VB.Label M3 
         AutoSize        =   -1  'True
         Caption         =   "3 . 1학년 데이터베이스 다운로드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   15
         Top             =   3720
         Width           =   7410
      End
      Begin VB.Label M2 
         AutoSize        =   -1  'True
         Caption         =   "2 . 관리자 관리 (SuperAdmin Only)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   14
         Top             =   3120
         Width           =   8160
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3480
         TabIndex        =   13
         Top             =   1560
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "관리자 정보 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   480
         TabIndex        =   12
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label M1 
         AutoSize        =   -1  'True
         Caption         =   "1 . 학생 관리"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   2985
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "관리자 모드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   5130
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   840
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   1680
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "3학년 업데이트 사항이 있습니다!"
      Height          =   180
      Left            =   0
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "2학년 업데이트 사항이 있습니다!"
      Height          =   180
      Left            =   0
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1학년 업데이트 사항이 있습니다!"
      Height          =   180
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   60
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8160
      TabIndex        =   23
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "최근 인증한 학번 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   -2880
      TabIndex        =   8
      Top             =   3360
      Width           =   7980
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5280
      TabIndex        =   7
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Made By..."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7560
      TabIndex        =   6
      Top             =   6240
      Width           =   2190
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5760
      TabIndex        =   5
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "학번을 입력해주세요."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1440
      TabIndex        =   4
      Top             =   5040
      Width           =   9300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3240
      TabIndex        =   3
      Top             =   6120
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5280
      TabIndex        =   2
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "입력하신 학번 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   -1560
      TabIndex        =   1
      Top             =   2280
      Width           =   6705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "학생 정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserData As String '석식하는 학생들
Public UserData2 As String '모든학생 바코드 정보
Public AdminData As String
Public StatusA As String
Public temp As String
Public PersonalData As String
Public NPersonalData As String
Public winhttp As New winhttp.WinHttpRequest
Public DoingA As Boolean

Public AdminMode As Long
Public AdminData1 As String
Public AdminData2 As String
Public Start As Boolean
Public TempKey As String
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Public Function Clean(str As String) As String
For i = 0 To UBound(Split(str, "" & vbCrLf & "" & vbCrLf & ""))
 Clean = Replace(str, "" & vbCrLf & "" & vbCrLf & "", "")
 DoEvents
 Next i
End Function
Public Function Addlog(str As String) As String
Form4.List1.AddItem "" & time & " >> " & str & "", 0
End Function
Public Function DataLoad() As String
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
UserData = B.readall
 B.Close
  Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase2.db", 1, True)
UserData2 = B.readall
 B.Close
   Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
AdminData = B.readall
 B.Close
End Function
Public Function DeleteGrade(Grade As String) As String
Dim Ar, B, C As String, D As Long
On Error Resume Next
Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close
For i = 0 To UBound(Split(C, vbCrLf)) - 1
If Left(Split(C, vbCrLf)(i), "1") = Grade Then C = Replace(C, Split(C, vbCrLf)(i), ""): D = D + 1
DoEvents
Next i
C = Clean("" & C & "")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 B.Close
 DeleteGrade = D
End Function
Public Function DownloadGrade(Grade As String) As String

Dim Ar, B, C As String
 Set Ar = CreateObject("scripting.FileSystemObject")
On Error Resume Next
winhttp.Open "GET", "http://npsoft.kr/taejang/" & Grade & "grade.txt"
winhttp.Send
'MsgBox StrConv(winhttp.ResponseText, vbUnicode)
'Exit Function
DeleteGrade (Grade)
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close
DownloadGrade = UBound(Split(npd(Replace(winhttp.ResponseText, vbCrLf, "")), "/")) / 2
C = "" & C & "" & vbCrLf & "" & npd(Replace(winhttp.ResponseText, vbCrLf, "")) & ""
Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
B.Write (C)
B.Close
 MsgBox "" & UBound(Split(npd(Replace(winhttp.ResponseText, vbCrLf, "")), "/")) / 2 & "개의 데이터베이스 다운 완료.", vbInformation, "Success"
End Function
Public Function Inj(Infor As String, TypeA As String) As String
Label8.Caption = Infor
If InStr(UserData, "" & Infor & "/") Then
PersonalData = "" & Infor & "" & Split(Split(UserData, Infor)(1), vbCrLf)(0) & ""
If TypeA = "K" Then
If InStr(UserData2, Infor) Then
If Split(Split(Split(UserData2, Infor)(1), vbCrLf)(0), "/")(2) = "1" Then
StatusA = "2"
Label4.Caption = "인증 실패 (Error Code : A03)"
Label4.Visible = True
DataLoad
TempKey = ""
Label3.Caption = TempKey
Exit Function
End If
End If
End If
If InStr(Split(PersonalData, "/")(2), Date) Then
StatusA = "2"
Label4.Caption = "인증실패 (Error Code : A01)"
Label4.Visible = True
Else
NPersonalData = Replace(PersonalData, Split(PersonalData, "/")(2), Date)
UserData = Replace(UserData, PersonalData, NPersonalData)
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (UserData)
 B.Close
Addlog ("" & Split(PersonalData, "/")(1) & "(" & Split(PersonalData, "/")(0) & ") 학생 석식 완료")
StatusA = "1"
Label4.Caption = "인증성공"
Label4.Visible = True
End If
Else
StatusA = "2"
Label4.Caption = "인증실패 (Error Code : A02)"
Label4.Visible = True

End If
DataLoad
TempKey = ""
Label3.Caption = TempKey
End Function

Private Sub Command2_Click()

End Sub




Private Sub Check1_Click()
If Check1.Value = "1" Then
Timer7.Enabled = True
Else
Timer7.Enabled = False
End If
End Sub

Private Sub Form_Load()

DataLoad
hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf LowLevelKeyboardProc, App.hInstance, 0)

'Label10.Caption = "(사전 임시사용 기간입니다.6월쯤부터 모든 학생들이 사용가능합니다.)" & vbCrLf & "바코드 기능은 등록한 학생들에 한해서만 사용 가능합니다." & vbCrLf & "(등록기간 : 21일~23일 점심시간 두드림실, 학생증 지참)"
Label7.Caption = "Made By" & vbCrLf & "- 15대 학생회 3학년 부회장" & vbCrLf & "- POT 동아리 1대,2대 기장" & vbCrLf & "조경현"
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnhookWindowsHookEx hHook
End
End Sub

Private Sub Label3_Change()
If AdminMode = "0" Then
If StatusA = "0" Then
Label4.Visible = False
ElseIf StatusA = "1" Then
Label4.Visible = True
StatusA = "0"
ElseIf StatusA = "2" Then
StatusA = "0"
End If

If Len(TempKey) < "3" Then
Form3.List1.Clear
Form3.Label2.Caption = Form3.List1.ListCount
ElseIf Len(TempKey) = "3" Then

On Error Resume Next
Form3.List1.Clear

Dim Ar, B, C, D
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)



For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then
If Left(D, 3) = TempKey Then Form3.List1.AddItem D: D = ""
Form3.Label2.Caption = Form3.List1.ListCount
End If
DoEvents
Next i
End If
Label3.Caption = TempKey
Else
M1.ForeColor = RGB(0, 0, 0)
M2.ForeColor = RGB(0, 0, 0)
M3.ForeColor = RGB(0, 0, 0)
M4.ForeColor = RGB(0, 0, 0)
M5.ForeColor = RGB(0, 0, 0)
M6.ForeColor = RGB(0, 0, 0)
M7.ForeColor = RGB(0, 0, 0)
M8.ForeColor = RGB(0, 0, 0)
M9.ForeColor = RGB(0, 0, 0)
M0.ForeColor = RGB(0, 0, 0)
If Right(TempKey, 1) = "1" Then
M1.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "2" Then
M2.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "3" Then
M3.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "4" Then
M4.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "5" Then
M5.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "6" Then
M6.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "7" Then
M7.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "8" Then
M8.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "9" Then
M9.ForeColor = RGB(255, 0, 0)
ElseIf Right(TempKey, 1) = "0" Then
M0.ForeColor = RGB(255, 0, 0)
End If
End If
End Sub

Private Sub Timer1_Timer()
Label1.Left = Me.Width / 2 - Label1.Width / 2
Frame1.Left = Me.Width / 2 - Frame1.Width / 2
Frame1.Top = Me.Height / 2 - Frame1.Height / 2
Frame2.Left = Me.Width / 2 - Frame2.Width / 2
Frame2.Top = Me.Height / 2 - Frame2.Height / 2
Label5.Left = Me.Width / 2 - Label5.Width / 2
Label2.Left = Label1.Left - Label1.Width / 2
Label3.Left = Label2.Left + Label2.Width + 40
Label9.Left = Label2.Left + Label2.Width - Label9.Width
Label8.Left = Label9.Left + Label9.Width + 40
Label4.Left = Me.Width / 2 - Label4.Width / 2
Label6.Left = Label3.Left + Label3.Width + 100
Label5.Left = Me.Width / 2 - Label5.Width / 2
Label7.Left = Me.Width - Label7.Width - 80
Label7.Top = Me.Height - Label7.Height - 500
Label11.Left = Label1.Left + Label1.Width + 100
'Label10.Top = Me.Height - Label10.Height - 500
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Len(Label3.Caption) <> 5 Then
Label5.Caption = "학번을 입력해주세요."
Label6.Caption = ""
Else
Dim Ar, B, C
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
If InStr(C, Label3.Caption) Then
Label6.Caption = Split(Split(C, Label3.Caption)(1), "/")(1)
Label5.Caption = "입력하신 학번이 맞으면" & vbCrLf & "엔터를 눌러주세요."
End If
End If
End Sub

Private Sub Timer3_Timer()
If AdminMode = "1" Or AdminMode = "2" Then
Frame2.Visible = False
Frame1.Visible = True
Me.BackColor = &H0&
ElseIf AdminMode = "0" Then
Frame1.Visible = False
Frame2.Visible = False
Me.BackColor = &H8000000F
ElseIf AdminMode = "3" Then
Frame1.Visible = False
Frame2.Visible = True

Me.BackColor = &H0&
End If
End Sub


Private Sub Timer4_Timer()
On Error Resume Next
If GetAsyncKeyState(vbKeyReturn) <> 0 Then
Dim TempLV As String
        TempKey = Replace(Replace(Replace(Replace(Replace(TempKey, " ", ""), "+", ""), "-", ""), "*", ""), "/", "")
If AdminMode = "0" Then '일반 모드
If TempKey = "" Or Len(TempKey) <= "4" Or Len(TempKey) = "6" Or Len(TempKey) >= "8" Then TempKey = "": Label3.Caption = TempKey: Exit Sub
If InStr(AdminData, "" & TempKey & "/") Then
AdminMode = Split(Split(AdminData, "" & TempKey & "")(1), "/")(2)
If Split(Split(AdminData, "" & TempKey & "")(1), "/")(2) = "2" Then
TempLV = "Super Admin"
Else
TempLV = "Normal Admin"
End If
Label13.Caption = "" & Split(Split(AdminData, "" & TempKey & "")(1), "/")(1) & " (" & TempLV & ")"
AdminData1 = TempKey
TempKey = ""
ElseIf Len(TempKey) = "5" Then
Call Inj(TempKey, "K")
ElseIf Len(TempKey) = "7" Then
If InStr(UserData2, TempKey) Then
TempKey = Split(Split(UserData2, TempKey)(1), "/")(1)
Else
TempKey = "00000"
End If
Call Inj(TempKey, "B")
Else '잘못된경우

Exit Sub
End If
ElseIf AdminMode = "3" Then '급식 Close

If TempKey = "" Then Exit Sub
If InStr(AdminData, "" & TempKey & "/") Then
AdminMode = Split(AdminData, "/")(2)
If Split(AdminData, "/")(2) = "2" Then
TempLV = "Super Admin"
Else
TempLV = "Normal Admin"
End If
Label13.Caption = "" & Split(AdminData, "/")(1) & " (" & TempLV & ")"
AdminData1 = TempKey
End If
TempKey = ""

ElseIf AdminMode = "1" Or AdminMode = "2" Then '관리자 모드
If Right(TempKey, 1) = "0" Then
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "1" Then
DoingA = True
Form2.Show
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "2" Then
If AdminMode = "2" Then
DoingA = True
Form6.Show
Else
MsgBox "해당 메뉴는 SuperAdmin 만 입장이 가능합니다." & vbCrLf & "SuperAdmin 에게 권한 상승을 요청하세요.", vbCritical, "Error"
End If
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "3" Then
DownloadGrade ("1")
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "4" Then
DownloadGrade ("2")
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "5" Then
DownloadGrade ("3")
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "6" Then
Form4.Show
DoingA = True
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "7" Then
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "8" Then
AdminMode = "3"
TempKey = ""
Label3.Caption = TempKey
ElseIf Right(TempKey, 1) = "9" Then
AdminMode = "0"
TempKey = ""
Label3.Caption = TempKey
If InStr(UserData2, AdminData1) Then
TempKey = Split(Split(UserData2, AdminData1)(1), "/")(1)
Else
TempKey = "00000"
End If
Call Inj(TempKey, "A")


End If
TempKey = ""
End If
If AdminMode <> "0" Then TempKey = ""
Form4.Label3.Caption = UBound(Split(UserData, Label4.Caption))
End If
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
UserData = B.readall

Label11.Caption = "(" & UBound(Split(UserData, Date)) & "명)"
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
winhttp.Open "GET", "http://npsoft.kr/taejang/updategrade1.txt"
winhttp.Send
If InStr(winhttp.ResponseText, "1") Then
Label15.Visible = True
Else
Label15.Visible = False
End If
winhttp.Open "GET", "http://npsoft.kr/taejang/updategrade2.txt"
winhttp.Send
If InStr(winhttp.ResponseText, "1") Then
Label16.Visible = True
Else
Label16.Visible = False
End If
winhttp.Open "GET", "http://npsoft.kr/taejang/updategrade3.txt"
winhttp.Send
If InStr(winhttp.ResponseText, "1") Then
Label17.Visible = True
Else
Label17.Visible = False
End If


End Sub

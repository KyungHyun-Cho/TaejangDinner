VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  '���� ����
   Caption         =   "�л� ���� ��"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6090
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command14 
      Caption         =   "������ ����"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   6120
      TabIndex        =   34
      Text            =   "	"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "��ü ���ĺ���"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "��ü ���ĿϷ�"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "���� ���ĺ���"
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���� ���ĿϷ�"
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "�˻�"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "�˻�"
         Height          =   615
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  '���
         Height          =   270
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '���
         Height          =   270
         Left            =   840
         MaxLength       =   5
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�̸�"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�й�"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   4440
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "���� ���"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   720
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   975
         Left            =   1880
         TabIndex        =   16
         Top             =   240
         Width           =   585
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   720
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   720
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "���� :"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�� :"
         Height          =   180
         Left            =   300
         TabIndex        =   12
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�г� :"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ü ����"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���� ����"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "�л� �߰�"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
      Begin MSComDlg.CommonDialog cd 
         Left            =   1920
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command12 
         Caption         =   "���� ���ε�"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '���
         Height          =   270
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   270
         Left            =   720
         MaxLength       =   5
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  '��� ����
         Caption         =   "�߰��� Enter"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�̸� :"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�й� :"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List"
      Height          =   5775
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command13 
         Caption         =   "��� ����"
         Height          =   300
         Left            =   120
         TabIndex        =   35
         Top             =   5400
         Width           =   3015
      End
      Begin VB.CommandButton Command7 
         Caption         =   "��ü ���� ����"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "��ü ����"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Appearance      =   0  '���
         Height          =   4440
         Left            =   120
         Style           =   1  'Ȯ�ζ�
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   180
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�� �л� �� :"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Access1 As Boolean '�л� ��ȸ
Public Access2 As Boolean '�л� �߰�
Public Access3 As Boolean '�л� �˻�
Public Access4 As Boolean '�л� ����
Public Access5 As Boolean '��ü ���� �Ϸ�/���� ����
Public Access6 As Boolean
Public Access7 As Boolean
Public Access8 As Boolean


Private Sub Command1_Click()
On Error Resume Next
List1.Clear


Dim Ar, B, C, D
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)


If Combo1.Text = "��ü�г�" Then
If Combo2.Text = "��ü��" Then
For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then List1.AddItem D: D = ""
DoEvents
Next i
Else
For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then
If Mid(D, 2, 2) = Combo2.Text Then List1.AddItem D: D = ""
End If
DoEvents
Next i
End If
Else
If Combo2.Text = "��ü��" Then
For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then
If Left(D, 1) = Combo1.Text Then List1.AddItem D: D = ""
End If
DoEvents
Next i
Else
For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then
If Left(D, 3) = "" & Combo1.Text & "" & Combo2.Text & "" Then List1.AddItem D: D = ""
End If
DoEvents
Next i
End If
End If
If Combo3.Text = "��ü" Then
ElseIf Combo3.Text = "����" Then
For i = List1.ListCount - 1 To 0 Step -1
If Split(List1.List(i), "/")(2) = Form1.Label4.Caption Then
List1.RemoveItem (i)
End If
DoEvents
Next i
ElseIf Combo3.Text = "�Ұ�" Then
For i = List1.ListCount - 1 To 0 Step -1
If Split(List1.List(i), "/")(2) <> Form1.Label4.Caption Then
List1.RemoveItem (i)
End If
DoEvents
Next i
End If
Label7.Caption = List1.ListCount
MsgBox "�˻��� �Ϸ�Ǿ����ϴ�.", vbInformation, "Success"
End Sub

Private Sub Command12_Click()
cd.ShowOpen
On Error Resume Next
If cd.FileName = "" Then Exit Sub
Dim Ar, B, C, D, E, I0, I1, I2, I3
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile(cd.FileName, 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile(cd.FileName, 1, True)
For i = 0 To C - 1
D = B.ReadLine

If UBound(Split(D, Text5.Text)) >= 2 Then
If UBound(Split(D, Text5.Text)) = "2" Then
If I0 = "" Then I0 = InputBox("���ε� �Ͻ� ������ �г��� �Է����ּ���.", "�г� �Է�")
If Split(D, Text5.Text)(0) <> "" Then I1 = Split(D, Text5.Text)(0)
I2 = Split(D, Text5.Text)(1)
I3 = Split(D, Text5.Text)(2)
Else
I0 = Split(D, Text5.Text)(0)
I1 = Split(D, Text5.Text)(1)
I2 = Split(D, Text5.Text)(2)
I3 = Split(D, Text5.Text)(3)
End If


If Len(I1) = "1" Then I1 = "0" & I1 & ""
If Len(I2) = "1" Then I2 = "0" & I2 & ""
D = "" & I0 & "" & I1 & "" & I2 & "/" & I3 & "/1990-00-00"
D = Replace(D, " ", "")
E = "" & E & "" & vbCrLf & "" & D & ""
List1.AddItem D: D = ""
Label7.Caption = List1.ListCount
End If
DoEvents
Next i
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 8, True)
B.WriteLine "" & E & "" & vbCrLf & ""
MsgBox "���ε尡 �Ϸ�Ǿ����ϴ�.", vbInformation, "Success"
End Sub

Private Sub Command10_Click()
If MsgBox("��ü �л����� ���� �Ϸ�(�Ұ�) ���·� ��ȯ�մϴ�.", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close

For i = 0 To List1.ListCount - 1
C = Replace(C, Split(List1.List(i), "/")(2), Form1.Label4.Caption)
DoEvents
Next i
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 B.Close
Call Command1_Click
End If

End Sub

Private Sub Command11_Click()
If MsgBox("��ü �л����� ���� ���� ���·� ��ȯ�մϴ�.", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close

For i = 0 To List1.ListCount - 1
C = Replace(C, Split(List1.List(i), "/")(2), "���� ����")
DoEvents
Next i
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 B.Close
Call Command1_Click
End If

End Sub

Private Sub Command13_Click()
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Statistic\StatisticSave(" & Combo1.Text & "" & Combo2.Text & "" & Combo3.Text & ")(" & Replace(Form1.Label4.Caption, "-", "") & "" & Replace(Replace(time, " ", ""), ":", "") & ").txt", True)
For i = 0 To List1.ListCount - 1
C = "" & C & "" & vbCrLf & "" & List1.List(i) & ""
DoEvents
Next i
C = "�� �л� �� : " & Label7.Caption & "" & vbCrLf & "" & C & ""
 B.Write (C)
B.Close
MsgBox "������ �Ϸ�Ǿ����ϴ�.", vbInformation, "����"
End Sub

Private Sub Command14_Click()
Form6.Show
End Sub

Private Sub Command2_Click()
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close
If List1.ListIndex = "-1" Then MsgBox "����Ʈ�� �������ּ���.", vbCritical, "Error": Exit Sub
If MsgBox("�����Ͻ� ����Ʈ�� �����˴ϴ�. ����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
If MsgBox("�����Ǹ� �ٽô� ������ �� �����ϴ�." & vbCrLf & "������ �����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "2�� Ȯ��") = vbYes Then
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then
C = Replace(C, "" & List1.List(i) & "", "")
List1.RemoveItem i
End If
DoEvents
Next i
End If
End If
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 B.Close
Label7.Caption = List1.ListCount
End Sub

Private Sub Command3_Click()
Dim Ar, B, C



If MsgBox("��ü ����Ʈ�� �����˴ϴ�. ����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
If MsgBox("�����Ǹ� �ٽô� ������ �� �����ϴ�." & vbCrLf & "������ �����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "2�� Ȯ��") = vbYes Then

 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
List1.Clear
B.Write (vbCrLf)
B.Close
End If
End If

Label7.Caption = List1.ListCount
End Sub

Private Sub Command4_Click()
List1.ListIndex = "4"
End Sub

Private Sub Command5_Click()
Dim K As Long
K = "0"
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
DoEvents
Next i
If Option1.Value = True Then
If Len(Text3.Text) <> 5 Then MsgBox "�й��� Ȯ�����ּ���.", vbCritical, "Error": Exit Sub
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), "" & Text3.Text & "/") Then
List1.Selected(i) = True
List1.ListIndex = i
K = K + 1
End If
DoEvents
Next i
Else
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), "/" & Text4.Text & "/") Then
List1.Selected(i) = True
List1.ListIndex = i
K = K + 1
End If
DoEvents
Next i
End If
MsgBox "�� " & K & " ���� �л��� �˻��Ǿ����ϴ�.", vbInformation, "ã��"
End Sub

Private Sub Command6_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
DoEvents
Next i
End Sub

Private Sub Command7_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
DoEvents
Next i
End Sub

Private Sub Command8_Click()
If MsgBox("������ �л����� ���� �Ϸ�(�Ұ�) ���·� ��ȯ�մϴ�.", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then C = Replace(C, List1.List(i), "" & Split(List1.List(i), "/")(0) & "/" & Split(List1.List(i), "/")(1) & "/" & Form1.Label4.Caption & "")
DoEvents
Next i
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 
 B.Close
Call Command1_Click


End If

End Sub

Private Sub Command9_Click()
If MsgBox("������ �л����� ���� �Ϸ�(�Ұ�) ���·� ��ȯ�մϴ�.", vbInformation + vbYesNo, "Ȯ��") = vbYes Then
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = B.readall
B.Close

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then C = Replace(C, List1.List(i), "" & Split(List1.List(i), "/")(0) & "/" & Split(List1.List(i), "/")(1) & "/���� ����")
DoEvents
Next i
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (C)
 B.Close
Call Command1_Click
End If


End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Ar, B, C, D
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)


For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then List1.AddItem D: D = ""
DoEvents
Next i
Label7.Caption = List1.ListCount

Combo1.AddItem "��ü�г�"
Combo2.AddItem "��ü��"
Combo3.AddItem "��ü"
Combo3.AddItem "����"
Combo3.AddItem "�Ұ�"
For i = 1 To 3
Combo1.AddItem i
DoEvents
Next i
For i = 1 To 16
If Len(i) = "1" Then
Combo2.AddItem "0" & i & ""
Else
Combo2.AddItem i
End If
DoEvents
Next i
B.Close
Combo1.ListIndex = "0"
Combo2.ListIndex = "0"
Combo3.ListIndex = "0"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
Form1.UserData = B.readall
 B.Close
Form7.DoingA = False
Unload Me
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = "5" Then Text2.SetFocus
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List1.AddItem "" & Text1.Text & "/" & Text2.Text & "/1990-00-00"
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 8, True)
B.WriteLine "" & Text1.Text & "/" & Text2.Text & "/1990-00-00"
 B.Close
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
Label7.Caption = List1.ListCount
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command5_Click
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command5_Click
End Sub

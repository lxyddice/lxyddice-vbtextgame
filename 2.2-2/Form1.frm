VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "��ҳ��v2.2-���ĵĴ���"
   ClientHeight    =   7035
   ClientLeft      =   7425
   ClientTop       =   2895
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   13170
   Begin VB.CommandButton Command21 
      Caption         =   "��Ϸ��Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   44
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   18
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "�������III��3����5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   42
      Top             =   8040
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar ProgressBar5 
      Height          =   135
      Left            =   2760
      TabIndex        =   40
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   135
      Left            =   6120
      TabIndex        =   37
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command20 
      Caption         =   "������ս��¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   36
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2760
      TabIndex        =   35
      Top             =   5190
      Width           =   2415
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   495
      Left            =   5640
      TabIndex        =   34
      Top             =   6240
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   135
      Left            =   480
      TabIndex        =   31
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer6 
      Interval        =   2500
      Left            =   12840
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   13080
      Top             =   0
   End
   Begin VB.CommandButton Command13 
      Caption         =   "��ʼ��ս"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   25
      Top             =   4920
      Width           =   5535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "�з�״̬"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   24
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12480
      Top             =   0
   End
   Begin VB.CommandButton Command11 
      Caption         =   "�ҷ�״̬"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   23
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   11880
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   11280
      Top             =   0
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   20
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   19
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   17
      Top             =   3480
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   480
      TabIndex        =   12
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10680
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   16
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   15
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   14
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ˢ�¼��ܣ�ʣ�ࣺ20��1��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   22
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0007
      TabIndex        =   10
      Top             =   2040
      Width           =   8655
   End
   Begin VB.CommandButton Command18 
      Caption         =   "������ǿҩ����5��:3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   33
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command17 
      Caption         =   "ҩˮ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   32
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "��ʦ�������ָ�ҩ����4����7"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   41
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Сƿħ��ҩ����2����8"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   43
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label17 
      Caption         =   "����ֵ��0000/0000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "�з�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "ħ��ֵ��0000/0000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "�����浵"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   29
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "�����¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "�˳���Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "�� ѡ �� �� Ϸ �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   26
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label10 
      Caption         =   "Ԫ�����ˣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "�������ԣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Ԫ�����ˣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "�������ԣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "�з�����ֵ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "�ҷ�����ֵ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wfsmsx As Long
Dim wfsm As Long
Dim dfsmsx As Long
Dim dfsm As Long
Dim wfgj As Long
Dim dfgj As Long
Dim wffy, wffysx As Long
Dim dffy As Long
Dim wffskx As Long
Dim dffskx As Long
Dim wfqsss As Long
Dim dfqsss As Long
Dim jn1 As Long
Dim jn2 As Long
Dim jn3 As Long
Dim jn4 As Long
Dim wfwlsh As Long
Dim wffssh As Long
Dim wfzzsh As Long
Dim dfwlsh As Long
Dim dffssh As Long
Dim dfzzsh As Long
Dim jnsxcs As Long
Dim dfgjfs As Long
Dim wfwlhd As Long
Dim wffshd As Long
Dim sjz1 As Long
Dim difficulty As Long
Dim wfcs As Long
Dim dfcs As Long
Dim bczzmc As Long
Dim sfbcg As Long
Dim dfsjss As Long
Dim wfcszt As Long
Dim wfbjl As Long
Dim dfbjl As Long
Dim wfbjsh, dfbjsh As Long
Dim sjz2 As Long
Dim fn As Integer, i As Integer
Dim wfmfz, wfmfsx As Long
Dim ysk As Long
Dim fsjqyj As Long
Dim wfsmws, wfgjws, wffyws, wffkws As Long
Dim drj1 As Double
Dim dfhd, dfhdsx As Long
Dim wfkjzsx, wfkjz As Long
Dim dsjsmhfyj, jjss3 As Long
Dim wfjfy, wfyjjfy As Long
Private Sub Command1_Click()

If dfcs > 0 Then
dfcs = dfcs - 1
End If



wfwlsh = 0
wffssh = 0

If jn1 = 1 Then
If wfmfz >= 4 Then
wfmfz = wfmfz - 4
wfwlsh = wfgj * 1.2 * 3 - (3 * dffy)
List1.AddItem "��ʹ�ü��ܣ��̿�������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn1 = 2 Then
If wfmfz >= 5 Then
wfmfz = wfmfz - 5
wfwlsh = wfgj * 0.9 * 5 - (5 * dffy)
List1.AddItem "��ʹ�ü��ܣ��̿�������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn1 = 3 Then
If wfmfz >= 7 Then
wfmfz = wfmfz - 7
wfwlsh = wfgj * 3.5 - (1 * dffy)
List1.AddItem "��ʹ�ü��ܣ���ɱ"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn1 = 4 Then
If wfmfz >= 9 Then
wfmfz = wfmfz - 9
wfwlsh = wfgj * 3 - (1 * dffy)
wfwlsh = wfwlsh + 3000
List1.AddItem "��ʹ�ü��ܣ���������������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If wfzzsh < 0 Then
List1.AddItem "�����˺�С�ڵз�������"
End If

If fsjqyj > 0 Then
wffssh = wffssh * 1.5
fsjqyj = fsjqyj - 1
End If

If dfhd > wffssh Then
If dfhd >= 3 * wffssh Then
dfhd = dfhd - wffssh * 3
wffssh = 0
Else
wffssh = wffssh - dfhd / 3
dfhd = 0
End If
End If

If dfhd >= wfwlsh Then
dfhd = dfhd - wfwlsh
wfwlsh = 0
Else
wfwlsh = wfwlsh - dfhd
dfhd = 0
End If

wfzzsh = wfwlsh + wffssh

sjz2 = Int(Rnd * (1000 - 0 + 1)) + 0
If sjz2 < wfbjl Then
wfzzsh = wfzzsh * (1 + wfbjsh / 1000)
List1.AddItem "�㴥���˱����������" & wfzzsh & "���˺�"
Else
List1.AddItem "�������" & wfzzsh & "���˺�"
End If




Timer3.Enabled = True

dfsm = dfsm - wfzzsh


Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command14.Visible = False
Command16.Visible = False
Command18.Visible = False

If dfcs > 0 Then
Timer2.Enabled = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Command14.Visible = True
Command16.Visible = True
Command18.Visible = True

List1.AddItem "�з�����˯�ˣ��㻹��" & dfcs & "�غϹ���"
Else
Timer2.Enabled = True
Timer4.Enabled = True
End If


End Sub

Private Sub Command10_Click()
jnsxcs = jnsxcs - 1
If jnsxcs < 0 Then
List1.AddItem "����ˢ�´������㣡"
Else
If wfmfz >= 1 Then
Timer2.Enabled = True
Command10.Caption = "ˢ�����м��ܣ�ʣ��" & jnsxcs & "��1��"
wfmfz = wfmfz - 1
Else
jnsxcs = jnsxcs + 1
List1.AddItem "ħ��ֵ���㣡"
End If
End If
End Sub
Private Sub Command11_Click()
List1.AddItem "�ҷ������ʣ�" & wfbjl / 10 & "%"
List1.AddItem "�ҷ������˺���" & wfbjsh / 10 & "%"
List1.AddItem "�ҷ������ܣ�" & wfwlhd
List1.AddItem "�ҷ��������ܣ�" & wffshd
List1.AddItem jjss3
End Sub

Private Sub Command12_Click()
List1.AddItem "�з������ʣ�" & dfbjl / 10 & "%"
List1.AddItem "�з������˺���" & dfbjsh / 10 & "%"
List1.AddItem "�з����ܣ�" & dfhd
End Sub

Private Sub Command13_Click()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = True
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command17.Visible = True
Command18.Visible = True
List1.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
ProgressBar1.Visible = True
ProgressBar2.Visible = True
ProgressBar3.Visible = True
ProgressBar4.Visible = True
ProgressBar5.Visible = True
Command13.Visible = False
List1.AddItem "������ս����Ҫ�Կ��ĵ����ǡ����ĵĴ�����"
If difficulty = 1 Then
wfsmsx = wfsmsx * 1.45
wfsm = wfsmsx
wfgj = wfgj * 1.35
wffy = wffy * 1.35
ProgressBar1.Max = wfsmsx
ProgressBar2.Max = dfsmsx
List1.AddItem "����С�ӣ��ҷ���λ����ֵ+45%���������������+30%"
End If
If difficulty = 3 Then
dfsmsx = dfsmsx * 1.4
dfsm = dfsmsx
dfgj = dfgj * 1.4
dffy = dffy * 1.4
ProgressBar1.Max = wfsmsx
ProgressBar2.Max = dfsmsx
List1.AddItem "����˹�䵶�����ез���λ�Ĺ�������������������+40%"
End If
End Sub



Private Sub Command14_Click()
If wfkjz >= 4 Then
dsjsmhfyj = 5
wfsm = wfsm + 7500
List1.AddItem "��ʦ�������ָ�ҩ����ʹ�óɹ���"
wfkjz = wfkjz - 4
Else
List1.AddItem "����ֵ���㣡"
End If
End Sub

Private Sub Command15_Click()
If wfkjz >= 3 Then
jjss3 = 5
List1.AddItem "�������III��ʹ�óɹ���"
wfkjz = wfkjz - 3
Else
List1.AddItem "����ֵ���㣡"
End If
End Sub

Private Sub Command16_Click()
If wfkjz >= 2 Then
wfmfz = wfmfz + 4
List1.AddItem "Сƿħ��ҩ����ʹ�óɹ���"
wfkjz = wfkjz - 2
Else
List1.AddItem "����ֵ���㣡"
End If
End Sub

Private Sub Command17_Click()
If ysk = 1 Then
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command18.Visible = False
Command14.Visible = False
Command17.Visible = False
ysk = 0
Else
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command11.Visible = False
Command12.Visible = False
Command18.Visible = True
Command14.Visible = True
Command17.Visible = True
ysk = 1
End If

End Sub

Private Sub Command18_Click()
If wfkjz >= 5 Then
fsjqyj = 3
List1.AddItem "������ǿҩ����ʹ�óɹ���"
wfkjz = wfkjz - 5
Else
List1.AddItem "����ֵ���㣡"
End If
End Sub

Private Sub Command2_Click()

If dfcs > 0 Then
dfcs = dfcs - 1
End If

wfwlsh = 0
wffssh = 0

If jn2 = 1 Then
If wfmfz >= 2 Then
wfmfz = wfmfz - 2
wffssh = wfgj * 1 * (1 - dffskx / 100)
List1.AddItem "��ʹ�ü��ܣ�������ͨ����"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn2 = 2 Then
If wfmfz >= 2 Then
wfmfz = wfmfz - 2
wffssh = wfgj * 0.8 * (1 - dffskx / 100)
dffy = dffy - 50
List1.AddItem "��ʹ�ü��ܣ��Ƽ׷���"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn2 = 3 Then
If wfmfz >= 2 Then
wfmfz = wfmfz - 2
wffssh = wfgj * 0.7 * (1 - dffskx / 100)
dfcs = dfcs + 1
List1.AddItem "��ʹ�ü��ܣ�����ج��"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn2 = 4 Then
If wfmfz >= 16 Then
wfmfz = wfmfz - 16
wffssh = wfgj * 8 * (1 - dffskx / 100)
dfcs = dfcs + 3
List1.AddItem "��ʹ�ü��ܣ����ѡ���ʥҫ��"

Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If wfzzsh < 0 Then
List1.AddItem "�����˺�С�ڵз�������"
End If

If fsjqyj > 0 Then
wffssh = wffssh * 1.5
fsjqyj = fsjqyj - 1
End If
If wfzzsh >= dfhd Then
wdzzsh = wfzzsh - dfhd
dfhd = 0
End If

If dfhd > wffssh Then
If dfhd >= 3 * wffssh Then
dfhd = dfhd - wffssh * 3
wffssh = 0
Else
wffssh = wffssh - dfhd / 3
dfhd = 0
End If
End If

If dfhd >= wfwlsh Then
dfhd = dfhd - wfwlsh
Else
wfwlsh = wfwlsh - dfhd
dfhd = 0
End If

wfzzsh = wfwlsh + wffssh

sjz2 = Int(Rnd * (1000 - 0 + 1)) + 0
If sjz2 < wfbjl Then
wfzzsh = wfzzsh * (1 + wfbjsh / 1000)
List1.AddItem "�㴥���˱����������" & wfzzsh & "���˺�"
Else
List1.AddItem "�������" & wfzzsh & "���˺�"
End If



Timer3.Enabled = True

dfsm = dfsm - wffssh

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command14.Visible = False
Command16.Visible = False
Command18.Visible = False

If dfcs > 0 Then
Timer2.Enabled = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Command14.Visible = True
Command16.Visible = True
Command18.Visible = True
List1.AddItem "�з�����˯�ˣ��㻹��" & dfcs & "�غϹ���"
Else
Timer2.Enabled = True
Timer4.Enabled = True
End If
End Sub

Private Sub Command20_Click()

If sjz2 = 1 Then
wffysx = Left(Text1.Text, 3)
wffy = wffysx

wffskx = Mid(Text1.Text, 4, 2)
Command13.Visible = True
Command20.Visible = False
Text1.Visible = False
End If

If sjz2 = 0 Then
If Text1.Text = "" Then
MsgBox "צ�ͣ��������ݰ���"
Else
wfsmws = Int(Text1.Text)
wfsm = Left(Text1.Text, 7)
wfgj = Right(Text1.Text, 4)
sjz2 = 1
Text1.Text = ""
End If

End If


End Sub

Private Sub Command21_Click()
frmAbout.Show
End Sub

Private Sub Command3_Click()

If dfcs > 0 Then
dfcs = dfcs - 1
End If

wfwlsh = 0
wffssh = 0

If jn3 = 1 Then
If wfmfz >= 2 Then
wfmfz = wfmfz - 2
wfwlhd = Int(Rnd * (1500 - 998 + 1)) + 998
List1.AddItem "��ʹ�ü��ܣ�������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn3 = 2 Then
If wfmfz >= 2 Then
wfmfz = wfmfz - 2
wffshd = Int(Rnd * (1500 - 998 + 1)) + 998

List1.AddItem "��ʹ�ü��ܣ���������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn3 = 3 Then
If wfmfz >= 4 Then
wfmfz = wfmfz - 4
wfwlhd = Int(Rnd * (2488 - 1288 + 1)) + 1288

List1.AddItem "��ʹ�ü��ܣ��߶�������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn3 = 4 Then
If wfmfz >= 10 Then
wfmfz = wfmfz - 10
wfwlhd = Int(Rnd * (2488 - 1288 + 1)) + 1288
dfcs = dfcs + 2
List1.AddItem "��ʹ�ü��ܣ��߶�������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If wfzzsh < 0 Then
List1.AddItem "�����˺�С�ڵз�������"
End If

If fsjqyj > 0 Then
wffssh = wffssh * 1.5
fsjqyj = fsjqyj - 1
End If

wfzzsh = wfwlsh + wffssh

If dfhd > wfzzsh Then
dfhd = dfhd - wfzzsh
wfzzsh = 0
End If
If wfzzsh >= dfhd Then
wdzzsh = wfzzsh - dfhd
dfhd = 0
End If

Timer3.Enabled = True

dfsm = dfsm - wfzzsh

List1.AddItem "��������" & wfwlhd & "���������"
List1.AddItem "��������" & wffshd & "��ķ�������"

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command14.Visible = False
Command16.Visible = False
Command18.Visible = False

If dfcs > 0 Then
Timer2.Enabled = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Command14.Visible = True
Command16.Visible = True
Command18.Visible = True
List1.AddItem "�з�����˯�ˣ��㻹��" & dfcs & "�غϹ���"
Else
Timer2.Enabled = True
Timer4.Enabled = True
End If
End Sub

Private Sub Command4_Click()

If dfcs > 0 Then
dfcs = dfcs - 1
End If

wfwlsh = 0
wffssh = 0

If jn4 = 1 Then
If wfmfz >= 1 Then
wfmfz = wfmfz - 1
dfqsss = dfqsss + 150 + wfgj * 0.2

List1.AddItem "��ʹ�ü��ܣ�ˮʴ��"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn4 = 2 Then
If wfmfz >= 1 Then
wfmfz = wfmfz - 1
dfqsss = dfqsss + wfgj * 0.3
wfwlsh = wfgj * 0.75 - dffy
List1.AddItem "��ʹ�ü��ܣ��ܼ�"

Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn4 = 3 Then
If wfmfz >= 1 Then
wfmfz = wfmfz - 1
dfqsss = dfqsss + 0.3 * wfgj
wffssh = wfgj * 0.25 * (1 - dffskx / 100)
List1.AddItem "��ʹ�ü��ܣ���ʴ֮ˮ"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn4 = 4 Then
If wfmfz >= 5 Then
wfmfsx = wfmfsx + 3
wfsm = wfsm + 1500
wfmfz = wfmfz - 5
wffssh = wfgj * 0.6 * (1 - dffskx / 100)
List1.AddItem "��ʹ�ü��ܣ����ף��"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If jn4 = 5 Then
If wfmfz >= 6 Then
wfsm = wfsm + 4514
wfmfz = wfmfz - 6

List1.AddItem "��ʹ�ü��ܣ����������"
Else
List1.AddItem "ħ��ֵ���㣬ʹ��ʧ�ܣ�"
End If
End If

If wfzzsh < 0 Then
List1.AddItem "�����˺�С�ڵз�������"
End If

If fsjqyj > 0 Then
wffssh = wffssh * 1.5
fsjqyj = fsjqyj - 1
End If
If wfzzsh >= dfhd Then
wdzzsh = wfzzsh - dfhd
dfhd = 0
End If

wfzzsh = wfwlsh + wffssh

sjz2 = Int(Rnd * (1000 - 0 + 1)) + 0
If sjz2 < wfbjl Then
wfzzsh = wfzzsh * (1 + wfbjsh / 1000)
List1.AddItem "�㴥���˱����������" & wfzzsh & "���˺�"
Else
List1.AddItem "�������" & wfzzsh & "���˺�"
End If

Timer3.Enabled = True

If dfhd > wfzzsh Then
dfhd = dfhd - wfzzsh
wfzzsh = 0
End If

dfsm = dfsm - wfzzsh


Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command14.Visible = False
Command16.Visible = False
Command18.Visible = False

If dfcs > 0 Then
Timer2.Enabled = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Command14.Visible = True
Command16.Visible = True
Command18.Visible = True
List1.AddItem "�з�����˯�ˣ��㻹��" & wfcs & "�غϹ���"
Else
Timer2.Enabled = True
Timer4.Enabled = True
End If
End Sub

Private Sub Command5_Click()
If jn1 = 1 Then

List1.AddItem "�̿����������Եз����3�ι�����120%�������˺�"
End If

If jn1 = 2 Then

List1.AddItem "�̿����������Եз����5�ι�����90%�������˺�"
End If

If jn1 = 3 Then

List1.AddItem "��ɱ���Եз���ɹ�����350%�������˺�"
End If

If jn1 = 4 Then

List1.AddItem "�����������������Եз���ɹ�����300%�������˺����������5000��ʵ�˺�"
End If

End Sub

Private Sub Command6_Click()
If jn2 = 1 Then

List1.AddItem "������ͨ�������Եз���ɹ�����100%�ķ����˺�"
End If

If jn2 = 2 Then

List1.AddItem "�Ƽ׷������Եз���ɹ�����80%�ķ����˺���������50������"
End If

If jn2 = 3 Then

List1.AddItem "����ج�Σ��Եз���ɹ�����70%�ķ����˺�����˯�з�1�غ�"
End If

If jn2 = 4 Then

List1.AddItem "���ѡ���ʥҫ�⣺�Եз���ɹ�����500%�ķ����˺�����˯�з�3�غ�"
End If



End Sub

Private Sub Command7_Click()
If jn3 = 1 Then

List1.AddItem "�����ܣ���ÿɵֵ������˺��Ļ���"
End If

If jn3 = 2 Then

List1.AddItem "�������ܣ���ÿɵֵ������˺��Ļ���"
End If

If jn3 = 3 Then

List1.AddItem "�߶������ܣ���ø߶�Ŀɵֵ������˺��Ļ���"
End If

If jn3 = 4 Then

List1.AddItem "�����������춯����(10)����ø߶�Ŀɵֵ������˺��Ļ��ܲ�ʯ������2�غ�"
End If
End Sub

Private Sub Command8_Click()
If jn4 = 1 Then
List1.AddItem "ˮʴ�ߣ���Ŀ�����150���빥����20%��ʴ����"
End If
If jn4 = 2 Then
List1.AddItem "�ܼף���Ŀ����ɹ�����30%��ʴ������75%�����˺�"
End If
If jn4 = 3 Then
List1.AddItem "��ʴ֮ˮ����Ŀ����ɹ�����30%��ʴ������25%�����˺�"
End If
If jn4 = 4 Then
List1.AddItem "���ף������Ŀ����ɹ�����60%�����˺����ظ�1500����ֵ������3��ħ��ֵ����"
End If
If jn4 = 5 Then
List1.AddItem "�����������ظ�4514������ֵ"
End If

End Sub

Private Sub Command9_Click()
List1.Clear
List1.AddItem wfcs
End Sub

Private Sub Form_Load()
Randomize
dfsm = 300000
wfsm = 200000
wfsmsx = 200000
dfsmsx = 300000
jnsxcs = 20
drjl = 0
dfhdsx = 100
dfhd = 0
wfkjzsx = 35
wfkjz = 30
wfyjjfy = 0
wfjfy = 0
sjz2 = 0

wfmfz = 20
wfmfsx = 20

sfgcg = 0
wfgj = Int(Rnd * (1288 - 1068 + 1)) + 1068 '����m>n����n~m�䣨����n��m��
dfgj = Int(Rnd * (1200 - 1000 + 1)) + 1000
wffy = Int(Rnd * (400 - 320 + 1)) + 320
dffy = Int(Rnd * (700 - 500 + 1)) + 500
wffskx = Int(Rnd * (20 - 0 + 1)) + 0
dffskx = Int(Rnd * (40 - 20 + 1)) + 20
wfbjl = Int(Rnd * (300 - 100 + 1)) + 100
dfbjl = Int(Rnd * (300 - 100 + 1)) + 100
wfbjsh = Int(Rnd * (888 - 150 + 1)) + 150
dfbjsh = Int(Rnd * (888 - 150 + 1)) + 150

bczzmc = Int(Rnd * (888888888 - 1 + 1)) + 1 '����m>n����n~m�䣨����n��m��

ProgressBar1.Max = wfsmsx
ProgressBar2.Max = dfsmsx
ProgressBar3.Max = wfmfsx

Command20.Visible = True
Text1.Visible = True

Label11.Visible = False


Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
Command11.Visible = False
Command12.Visible = False
Command13.Visible = False
Command17.Visible = False
Command18.Visible = False
Command14.Visible = False
List1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
Timer2.Enabled = True
End Sub

Private Sub Label12_Click()
End
End Sub

Private Sub Label13_Click()

If sfbcg > 0 Then

Kill App.Path & "\" & bczzmc & ".txt"
End If
sfbcg = sfbcg + 1
Dim fn As Integer, i As Integer
For i = 0 To List1.ListCount - 1
Open App.Path & "\" & bczzmc & ".txt" For Append As #1
Print #1, List1.List(i)
Close #1
Next
Close fn


End Sub

Private Sub Label14_Click()

If wfsm <= wfsmsx / 10 Then
wfsmws = Len(Label1.Caption) - 1
Else
wfsmws = Len(Label1.Caption)
End If
If wfsmws = 14 Then
wfsmws = "000000" & wfsm
End If
If wfsmws = 15 Then
wfsmws = "00000" & wfsm
End If
If wfsmws = 16 Then
wfsmws = "0000" & wfsm
End If
If wfsmws = 17 Then
wfsmws = "000" & wfsm
End If
If wfsmws = 18 Then
wfsmws = "00" & wfsm
End If
If wfsmws = 19 Then
wfsmws = "0" & wfsm
End If
If wfsmws = 20 Then
wfsmws = wfsm
End If


List1.AddItem wfsmws
'List1.AddItem wfsm & dfsm & wfsmsx & dfsmsx & wffy & dffy & wffskx & dffsxk & wfqsss & dfqsss & wfbjl & dfbjl & wfbjsh & dfbjsh & difficulty
End Sub

Private Sub Timer1_Timer()
List1.ListIndex = List1.ListCount - 1
Label1.Caption = "�ҷ�����ֵ��" & wfsm & "/" & wfsmsx
Label2.Caption = "�з�����ֵ��" & dfsm & "/" & dfsmsx
Label3.Caption = "��������" & wfgj
Label7.Caption = "��������" & dfgj
Label15.Caption = "ħ��ֵ��" & wfmfz & "/" & wfmfsx
Label17.Caption = "����ֵ��" & wfkjz & "/" & wfkjzsx

Label4.Caption = "��������" & wffy
Label8.Caption = "��������" & dffy
Label5.Caption = "�������ԣ�" & wffskx
Label9.Caption = "�������ԣ�" & dffskx
Label6.Caption = "Ԫ�����ˣ�" & wfqsss & "/1000"
Label10.Caption = "Ԫ�����ˣ�" & dfqsss & "/1000"



If wfjfy > 0 Then
wffy = wffysx - wfjfy
wfjfy = 0
End If

If wfqsss >= 1000 Then
wfsm = wfsm - 3250 + wffy
wffy = wffy - 100
wfqsss = 0

End If

If dfsjss >= 1000 Then
dfsm = dfsm - 3800
dfcs = dfcs + 2
dfsjss = 0
End If

If wfsm <= 0 Then
wfsm = 0
List1.AddItem wfsm & wfgj & wffy & wfbjl & wfbjsh
If sfbcg > 0 Then

Kill App.Path & "\" & bczzmc & ".txt"
End If
sfbcg = sfbcg + 1

For i = 0 To List1.ListCount - 1
Open App.Path & "\" & bczzmc & ".txt" For Append As #1
Print #1, List1.List(i)
Close #1
Next
Close fn
MsgBox "�����ˣ�"


End
End If
If dfsm <= 0 Then
dfsm = 0
List1.AddItem wfsm & wfgj & wffy & wfbjl & wfbjsh
If sfbcg > 0 Then

Kill App.Path & "\" & bczzmc & ".txt"
End If
sfbcg = sfbcg + 1

For i = 0 To List1.ListCount - 1
Open App.Path & "\" & bczzmc & ".txt" For Append As #1
Print #1, List1.List(i)
Close #1
Next
Close fn
MsgBox "��Ӯ�ˣ�"

If wffy < 0 Then
wffy = 0
End If

If dffy < 0 Then
dffy = 0
End If

End
End If

If wfmfz > wfmfsx Then
wfmfz = wfmfsx
End If

If wfsm > wfsmsx Then
wfsm = wfsmsx
End If

If dfsm > dfsmsx Then
dfsm = dfsmsx
End If

If wfkjz > wfkjzsx Then
wfkjz = wfkjzsx
End If

ProgressBar1.Max = wfsmsx
ProgressBar2.Max = dfsmsx
ProgressBar3.Max = wfmfsx
ProgressBar4.Max = dfhdsx
ProgressBar5.Max = wfkjzsx
ProgressBar1.Value = wfsm
ProgressBar2.Value = dfsm
ProgressBar3.Value = wfmfz
ProgressBar4.Value = dfhd
ProgressBar5.Value = wfkjz

If wfsm <= wfsmsx / 10 Then
Label1.Caption = "�ҷ�����ֵ��" & wfsm & "/" & wfsmsx & "��"
End If
If dfsm <= dfsmsx / 10 Then
Label2.Caption = "�з�����ֵ��" & dfsm & "/" & dfsmsx & "��"
End If
End Sub

Private Sub Timer2_Timer()
jn1 = Int(Rnd * (4 - 1 + 1)) + 1 '����m>n����n~m�䣨����n��m��
jn2 = Int(Rnd * (4 - 1 + 1)) + 1
jn3 = Int(Rnd * (4 - 1 + 1)) + 1
jn4 = Int(Rnd * (5 - 1 + 1)) + 1
If jn1 = 1 Then
Command1.Caption = "�̿�������(4)"
End If
If jn1 = 2 Then
Command1.Caption = "�̿�������(5)"
End If
If jn1 = 3 Then
Command1.Caption = "��ɱ(7)"
End If
If jn1 = 4 Then
Command1.Caption = "��������������(9)"
End If


If jn2 = 1 Then
Command2.Caption = "������ͨ����(2)"
End If
If jn2 = 2 Then
Command2.Caption = "�Ƽ׷���(2)"
End If
If jn2 = 3 Then
Command2.Caption = "����ج��(2)"
End If
If jn2 = 4 Then
Command2.Caption = "���ѡ���ʥҫ��(18)"
End If

If jn3 = 1 Then
Command3.Caption = "������(2)"
End If
If jn3 = 2 Then
Command3.Caption = "��������(2)"
End If
If jn3 = 3 Then
Command3.Caption = "�߶�������(4)"
End If
If jn3 = 4 Then
Command3.Caption = "�����������춯����(10)"
End If

If jn4 = 1 Then
Command4.Caption = "ˮʴ��(1)"
End If
If jn4 = 2 Then
Command4.Caption = "�ܼ�(1)"
End If
If jn4 = 3 Then
Command4.Caption = "��ʴ֮ˮ(1)"
End If
If jn4 = 4 Then
Command4.Caption = "���ף��(5)"
End If
If jn4 = 5 Then
Command4.Caption = "����nb(6)"
End If

If wffy < 0 Then
wffy = 0
End If

If dffy < 0 Then
dffy = 0
End If

Timer2.Enabled = False
End Sub
Private Sub Timer3_Timer()

dfsm = dfsm - wfzzsh

Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()


'If jjss3 > 0 Then
'wffy = wffysx * 1.5 - wfyjjfy
'jjss3 = jjss3 - 1
'Else
'wffy = wffy - wfyjjfy
'End If


dfwlsh = 0
dffssh = 0

If wfcs > 0 Then
wfcs = wfcs - 1
End If

dfgjfs = Int(Rnd * (210 - 1 + 1)) + 1
If dfgjfs > 1 And dfgjfs < 25 Then
List1.AddItem "�з�ʹ�ü��ܣ���ը"
dffssh = dfgj * 1.5 - (1 - wffskx / 100)
wfqsss = wfqsss + dfgj * 0.1
End If
If dfgjfs >= 25 And dfgjfs < 75 Then
List1.AddItem "�з�ʹ�ü��ܣ��ء���ͨ����"
dfwlsh = dfgj * 2.2 - wffy * 2
wfqsss = wfqsss + dfwlsh * 0.07 + 45
End If
If dfgjfs >= 75 And dfgjfs < 100 Then
List1.AddItem "�з�ʹ�ü��ܣ����"
dfwlsh = dfgj * 4 - wffy
wfqsss = wfqsss + 210
End If
If dfgjfs >= 100 And dfgjfs < 101 Then
List1.AddItem "�з�ʹ�ü��ܣ��ս�"
dfhdsx = dfsmsx * 0.1
dfhd = dfhdsx
End If
If dfgjfs >= 101 And dfgjfs < 125 Then
List1.AddItem "�з�ʹ�ü��ܣ�����"
dfsm = dfsm + dfsmsx * 0.05
wfqsss = wfqsss + dfgj * 0.3
End If

If dfgjfs >= 125 And dfgjfs < 211 Then
List1.AddItem "�з�ʹ�ü��ܣ���ͨ����"
dfwlsh = dfgj * 1 - wffy
wfqsss = wfqsss + dfgj * 0.02
End If


If wfwlhd > 0 Then
If wfwlhd > dfwlsh Then
wfwlhd = wfwlhd - dfwlsh
dfwlsh = 0
Else
dfwlsh = dfwlsh - wfwlhd
wfwlhd = 0
End If
End If

If wffshd > 0 Then
If wffshd > dffssh Then
wffshd = wffshd - dffssh
dffssh = 0
Else
dffssh = dffssh - wffshd
wffshd = 0
End If
End If

dfzzsh = dfwlsh + dffssh

sjz2 = Int(Rnd * (1000 - 0 + 1)) + 0

If sjz2 < dfbjl Then
dfzzsh = dfzzsh * (1 + dfbjsh / 1000)
List1.AddItem "�з������˱����������" & dfzzsh & "���˺�"
Else
List1.AddItem "�з������" & dfzzsh & "���˺�"
End If


If dsjsmhfyj > 0 Then
wfsm = wfsmsx * 0.012 + wfsm
dsjsmhfyj = dsjsmhfyj - 1
End If


wfsm = wfsm - dfzzsh

wfmfz = wfmfz + 3
If wfmfz > wfmfsx Then
wfmfz = wfmfsx
End If

Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True

Timer4.Enabled = False

If wfcs > 0 Then
If wfcszt = 1 Or wfcszt = 0 Then
List1.AddItem "�㱻��˯�ˣ��޷��ж���ʣ��" & wfcs & "�غ�"
End If
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Timer4.Enabled = True
End If
wfcszt = 0
End Sub

Private Sub Timer5_Timer()
If wfcs > 0 Then
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Timer5.Enabled = False
If wfcszt = 0 Then
List1.AddItem "�㱻��˯�ˣ��޷��ж���ʣ��" & wfcs & "�غ�"
wfcszt = 1
End If
Timer4.Enabled = True
End If
End Sub




Private Sub Timer6_Timer()
wfqsss = 0
Timer6.Enabled = False
End Sub

VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8250
   ClientLeft      =   2655
   ClientTop       =   3030
   ClientWidth     =   14025
   LinkTopic       =   "Form4"
   Picture         =   "frm_admin.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   14025
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   5760
      TabIndex        =   9
      Top             =   3840
      Width           =   4815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_admin.frx":0342
      Left            =   2400
      List            =   "frm_admin.frx":0352
      TabIndex        =   6
      Text            =   "Select Baud Rate"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Record #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Record #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Databits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Baud Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "GSM CONFIGURATION"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4050
      TabIndex        =   0
      Top             =   360
      Width           =   5595
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

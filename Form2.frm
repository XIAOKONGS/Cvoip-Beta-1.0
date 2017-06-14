VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Caption         =   "±ßÀÏÆÅ¼ÓÓÍ~ÄãÊÇ×î°ôßÕ"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form2"
   ScaleHeight     =   7155
   ScaleWidth      =   8700
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   7440
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   4440
      Width           =   3225
   End
   Begin VB.Label Label19 
      Caption         =   "Tahoma"
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   18
      Top             =   3600
      Width           =   3345
   End
   Begin VB.Label Label17 
      Caption         =   "·ÂËÎ"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ÐÂËÎÌå"
      Height          =   180
      Left            =   4440
      TabIndex        =   15
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label14 
      Caption         =   "Î¢ÈíÑÅºÚ"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "ÐÂËÎÌå"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   13
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   3045
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   11
      Top             =   9840
      Width           =   3495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   10920
      Width           =   3300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   10560
      Width           =   3015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   12600
      Width           =   3210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   7
      Top             =   11400
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   9960
      Width           =   3195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "·½ÕýÀ¼Í¤³¬Ï¸ºÚ¼òÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   8760
      Width           =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú ×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4800
      TabIndex        =   4
      Top             =   8880
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   8160
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "·½ÕýÀ¼Í¤³¬Ï¸ºÚ¼òÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   8040
      Width           =   3105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¡Ñ ¹§Ï²Äú×ªÒÆ³É¹¦À² ~"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   7080
      Width           =   3300
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

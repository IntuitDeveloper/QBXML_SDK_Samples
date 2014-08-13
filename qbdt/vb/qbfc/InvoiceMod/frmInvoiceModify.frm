VERSION 5.00
Begin VB.Form frmInvoiceModify 
   Caption         =   "Modify Invoice"
   ClientHeight    =   10035
   ClientLeft      =   3195
   ClientTop       =   795
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   9630
   Begin VB.CheckBox chkShow 
      Caption         =   "Show request"
      Height          =   255
      Left            =   6840
      TabIndex        =   177
      Top             =   10080
      Width           =   2055
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   29
      Left            =   4080
      TabIndex        =   170
      Top             =   8640
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   28
      Left            =   4080
      TabIndex        =   169
      Top             =   8400
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   27
      Left            =   4080
      TabIndex        =   168
      Top             =   8160
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   26
      Left            =   4080
      TabIndex        =   167
      Top             =   7920
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   25
      Left            =   4080
      TabIndex        =   166
      Top             =   7680
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   24
      Left            =   4080
      TabIndex        =   165
      Top             =   7440
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   23
      Left            =   4080
      TabIndex        =   164
      Top             =   7200
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   22
      Left            =   4080
      TabIndex        =   163
      Top             =   6960
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   21
      Left            =   4080
      TabIndex        =   162
      Top             =   6720
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   20
      Left            =   4080
      TabIndex        =   161
      Top             =   6480
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   19
      Left            =   7680
      TabIndex        =   160
      Top             =   8640
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   18
      Left            =   7680
      TabIndex        =   159
      Top             =   8400
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   17
      Left            =   7680
      TabIndex        =   158
      Top             =   8160
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   16
      Left            =   7680
      TabIndex        =   157
      Top             =   7920
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   15
      Left            =   7680
      TabIndex        =   156
      Top             =   7680
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   14
      Left            =   7680
      TabIndex        =   155
      Top             =   7440
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   13
      Left            =   7680
      TabIndex        =   154
      Top             =   7200
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   12
      Left            =   7680
      TabIndex        =   153
      Top             =   6960
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   11
      Left            =   7680
      TabIndex        =   152
      Top             =   6720
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   10
      Left            =   7680
      TabIndex        =   151
      Top             =   6480
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   150
      Top             =   8640
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   149
      Top             =   8400
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   148
      Top             =   8160
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   147
      Top             =   7920
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   146
      Top             =   7680
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   145
      Top             =   7440
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   144
      Top             =   7200
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   143
      Top             =   6960
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   142
      Top             =   6720
      Width           =   130
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   141
      Top             =   6480
      Width           =   130
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   140
      Top             =   8640
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   139
      Top             =   8640
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   138
      Top             =   8640
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   137
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   136
      Top             =   8640
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   135
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   134
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   133
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   132
      Top             =   8400
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   131
      Top             =   8400
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   130
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   129
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   128
      Top             =   8160
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   127
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   126
      Top             =   8160
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   125
      Top             =   8160
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   124
      Top             =   8160
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   123
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   122
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   121
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   120
      Top             =   7920
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   119
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   118
      Top             =   7920
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   117
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   116
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   115
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   114
      Top             =   7680
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   113
      Top             =   7680
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   112
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   111
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   110
      Top             =   7440
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   109
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   108
      Top             =   7440
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   107
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   106
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   105
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   104
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   103
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   102
      Top             =   7200
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   101
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   7200
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   99
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   96
      Top             =   6960
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   94
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   92
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   91
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   6720
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Return to Invoice Selection"
      Height          =   855
      Left            =   8280
      TabIndex        =   86
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoModify 
      Caption         =   "Modify Invoice"
      Height          =   855
      Left            =   6840
      TabIndex        =   85
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo All Invoice Changes"
      Height          =   855
      Left            =   5520
      TabIndex        =   84
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteUndelete 
      Caption         =   "Delete Line"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4080
      TabIndex        =   83
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLineBefore 
      Caption         =   "Add Line Before Current Line"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1440
      TabIndex        =   82
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLineAfter 
      Caption         =   "Add Line After Current Line"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      TabIndex        =   81
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditLine 
      Caption         =   "Edit Line"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   80
      Top             =   9120
      Width           =   1215
   End
   Begin VB.VScrollBar vscInvoiceLineScroll 
      Height          =   2415
      LargeChange     =   10
      Left            =   9240
      TabIndex        =   79
      Top             =   6480
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   6480
      Width           =   3495
   End
   Begin VB.TextBox txtItemFullName 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   6480
      Width           =   375
   End
   Begin VB.TextBox txtItemLineNumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditMemo 
      Caption         =   "Edit Memo"
      Height          =   255
      Left            =   8400
      TabIndex        =   72
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtMemo 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5760
      Width           =   7575
   End
   Begin VB.ComboBox cmbCustomerMsg 
      Height          =   315
      Left            =   1560
      TabIndex        =   69
      Top             =   5400
      Width           =   7815
   End
   Begin VB.ComboBox cmbCustTaxCode 
      Height          =   315
      Left            =   6840
      TabIndex        =   67
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox cmbItemSalesTax 
      Height          =   315
      Left            =   1320
      TabIndex        =   66
      Top             =   5040
      Width           =   3255
   End
   Begin VB.ComboBox cmbShipMethod 
      Height          =   315
      Left            =   3360
      TabIndex        =   65
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox cmbSalesRep 
      Height          =   315
      Left            =   960
      TabIndex        =   64
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtFOB 
      Height          =   285
      Left            =   8160
      TabIndex        =   59
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtShipDate 
      Height          =   285
      Left            =   6360
      TabIndex        =   58
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtDueDate 
      Height          =   285
      Left            =   4200
      TabIndex        =   57
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Height          =   285
      Left            =   1080
      TabIndex        =   56
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox cmbTerms 
      Height          =   315
      Left            =   5640
      TabIndex        =   51
      Top             =   3960
      Width           =   3855
   End
   Begin VB.ComboBox cmbARAccount 
      Height          =   315
      Left            =   1080
      TabIndex        =   49
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CheckBox chkToBePrinted 
      Caption         =   "To Be Printed"
      Height          =   255
      Left            =   7200
      TabIndex        =   47
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtShipAddr4 
      Height          =   285
      Left            =   5400
      TabIndex        =   46
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txtShipAddr3 
      Height          =   285
      Left            =   5400
      TabIndex        =   45
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox txtShipAddr2 
      Height          =   285
      Left            =   5400
      TabIndex        =   44
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtShipAddr1 
      Height          =   285
      Left            =   5400
      TabIndex        =   43
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtShipState 
      Height          =   285
      Left            =   8040
      TabIndex        =   42
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtShipCity 
      Height          =   285
      Left            =   5400
      TabIndex        =   41
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtShipPostalCode 
      Height          =   285
      Left            =   5880
      TabIndex        =   40
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtShipCountry 
      Height          =   285
      Left            =   7920
      TabIndex        =   39
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtBillCountry 
      Height          =   285
      Left            =   3120
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtBillPostalCode 
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtBillState 
      Height          =   285
      Left            =   3240
      TabIndex        =   26
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtBillCity 
      Height          =   285
      Left            =   600
      TabIndex        =   24
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtBillAddr4 
      Height          =   285
      Left            =   600
      TabIndex        =   23
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txtBillAddr3 
      Height          =   285
      Left            =   600
      TabIndex        =   22
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox txtBillAddr2 
      Height          =   285
      Left            =   600
      TabIndex        =   21
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtBillAddr1 
      Height          =   285
      Left            =   600
      TabIndex        =   20
      Top             =   1680
      Width           =   4095
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   2655
   End
   Begin VB.CheckBox chkPending 
      Caption         =   "Pending"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtTxnDate 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtRefNumber 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbCustomer 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtEditSequence 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtTxnID 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8470
      TabIndex        =   176
      Top             =   6240
      Width           =   665
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   175
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   174
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   173
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   172
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label26 
      Caption         =   "TxnLineID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   171
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label25 
      Caption         =   "Memo"
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "Customer Message"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label23 
      Caption         =   "Customer Sales Tax Code"
      Height          =   255
      Left            =   4800
      TabIndex        =   63
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "Item Sales Tax"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "Ship Method"
      Height          =   255
      Left            =   2280
      TabIndex        =   61
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "Sales Rep"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "FOB"
      Height          =   255
      Left            =   7680
      TabIndex        =   55
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "Ship Date"
      Height          =   255
      Left            =   5520
      TabIndex        =   54
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Due Date"
      Height          =   255
      Left            =   3360
      TabIndex        =   53
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "PO Number"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Terms"
      Height          =   255
      Left            =   5160
      TabIndex        =   50
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "AR Account"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Country"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   38
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Postal Code"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   37
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "State"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   36
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "City"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   35
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Addr4"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   34
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr3"
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   33
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr2"
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   32
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr1"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   31
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Country"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   29
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Postal Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "State"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   25
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "City"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Addr4"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr3"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Addr1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Ship Address"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Bill Address"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Class"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Date"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Invoice #"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Customer"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Edit Sequence"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "TxnID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmInvoiceModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

Dim strLineArray(200) As String
Dim strLastCommand As String

Dim booEditInvoiceLineFormLoaded As Boolean
Dim booCheckForChanges As Boolean
Dim booCheckForLineChanges As Boolean

Dim intNumInvoiceLines As Integer
Dim intTopDisplayedLine As Integer
Dim intHighlightedLine As Integer
Dim intActualHighlightedLine As Integer

Dim strTxnID As String
Dim strEditSequence As String
Dim strRefNumber As String
Dim strTxnDate As String
Dim intPending As Integer
Dim intToBePrinted As Integer
Dim strCustomer As String
Dim strClass As String
Dim strBillAddr1 As String
Dim strBillAddr2 As String
Dim strBillAddr3 As String
Dim strBillAddr4 As String
Dim strBillCity As String
Dim strBillState As String
Dim strBillPostalCode As String
Dim strBillCountry As String
Dim strShipAddr1 As String
Dim strShipAddr2 As String
Dim strShipAddr3 As String
Dim strShipAddr4 As String
Dim strShipCity As String
Dim strShipState As String
Dim strShipPostalCode As String
Dim strShipCountry As String
Dim strARAccount As String
Dim strTerms As String
Dim strPONumber As String
Dim strDueDate As String
Dim strShipDate As String
Dim strFOB As String
Dim strSalesRep As String
Dim strShipMethod As String
Dim strItemSalesTax As String
Dim strCustTaxCode As String
Dim strCustomerMsg As String
Dim strMemo As String

Dim booInvoiceBodyChanged As Boolean
Dim booBillAddressChanged As Boolean
Dim booShipAddressChanged As Boolean
Dim booInvoiceLinesChanged As Boolean
  
Private Sub Form_Load()
  
  strLastCommand = ""
  
  ClearForm

  frmPatience.Show
  FillComboBox cmbCustomer, "Customer", "FullName", "", False
  FillComboBox cmbClass, "Class", "FullName", "", False
  FillComboBox cmbARAccount, "Account", "FullName", "<AccountType>AccountsReceivable</AccountType>", False
  FillComboBox cmbTerms, "StandardTerms", "Name", "", False
  FillComboBox cmbSalesRep, "SalesRep", "FullName", "", False
  FillComboBox cmbShipMethod, "ShipMethod", "Name", "", False
  FillComboBox cmbItemSalesTax, "ItemSalesTax,ItemSalesTaxGroup", "Name", "", False
  FillComboBox cmbCustTaxCode, "SalesTaxCode", "Name", "", False
  FillComboBox cmbCustomerMsg, "CustomerMsg", "Name", "", False
  frmPatience.Hide
  
  intHighlightedLine = -1
  intActualHighlightedLine = -1

  LoadInvoiceModifyForm
  SaveOriginalValues
  LoadInvoiceLineArray strLineArray
  CountInvoiceLines
  DisplayInvoiceLines 1
  booEditInvoiceLineFormLoaded = False
  booCheckForChanges = False
  booCheckForLineChanges = False
End Sub


Private Sub Form_Activate()
  If strLastCommand = "ReturnToInvoiceSelection" Or _
     strLastCommand = "ModifyInvoice" Then
  
    ClearForm
    
    If intHighlightedLine > -1 Then
      UnHighlightLine intHighlightedLine
    End If
    
    intHighlightedLine = -1
    intActualHighlightedLine = -1

    LoadInvoiceModifyForm
    SaveOriginalValues
    LoadInvoiceLineArray strLineArray
    CountInvoiceLines
    If intNumInvoiceLines < 11 Or vscInvoiceLineScroll.Value = 1 Then
      DisplayInvoiceLines 1
    Else
      vscInvoiceLineScroll.Value = 1
    End If
    booEditInvoiceLineFormLoaded = False
    booCheckForChanges = False
    booCheckForLineChanges = False
  ElseIf strLastCommand = "EditLine" Then
    If GetInvoiceLineInfo <> strLineArray(intActualHighlightedLine) Then
      strLineArray(intActualHighlightedLine) = GetInvoiceLineInfo
      DisplayLine strLineArray(intActualHighlightedLine), intHighlightedLine
      booCheckForLineChanges = True
      End If
  ElseIf strLastCommand = "NewLineAfter" Then
    If InStr(1, GetInvoiceLineInfo, "<spliter>") = 0 Then Exit Sub
    
    booCheckForLineChanges = True
    InsertInvoiceLine GetInvoiceLineInfo, intActualHighlightedLine + 1
    
    If intNumInvoiceLines > 10 Then
      If intHighlightedLine = 9 Then
        vscInvoiceLineScroll.Value = intTopDisplayedLine + 1
      Else
        DisplayInvoiceLines intTopDisplayedLine
      End If
    Else
      DisplayInvoiceLines intTopDisplayedLine
    End If
    
  ElseIf strLastCommand = "NewLineBefore" Then
    If InStr(1, GetInvoiceLineInfo, "<spliter>") = 0 Then Exit Sub
    
    booCheckForLineChanges = True
    InsertInvoiceLine GetInvoiceLineInfo, intActualHighlightedLine
    
    Dim intOldHighlightedLine As Integer
    intOldHighlightedLine = intHighlightedLine
    UnHighlightLine intHighlightedLine
    intActualHighlightedLine = intActualHighlightedLine + 1
    
    If intNumInvoiceLines > 10 Then
      If intOldHighlightedLine = 9 Then
        vscInvoiceLineScroll.Value = intTopDisplayedLine + 1
      Else
        DisplayInvoiceLines intTopDisplayedLine
      End If
    Else
      DisplayInvoiceLines intTopDisplayedLine
    End If
    
  End If
End Sub

Private Sub cmdDoModify_Click()
  
  If Not InvoiceChanged Then
    If _
      MsgBox(vbCrLf & "This invoice appears to be unchanged" & vbCrLf & vbCrLf & _
             "Close this window and return to the invoice selection window?", _
             vbYesNo) = vbYes Then
      Unload frmInvoiceModify
    Else
      Exit Sub
    End If
  End If
  
  strLastCommand = "ModifyInvoice"
  frmModifying.Show
  ModifyInvoice InvoiceChangeString
  frmModifying.Hide
  If chkShow.Value = 1 Then
    Load frmShowRequest
    frmShowRequest.Show
    frmInvoiceModify.Hide
  Else
    frmInvoiceModify.Hide
  End If
End Sub

Private Sub cmdUndo_Click()
  UnHighlightLine intHighlightedLine
  intActualHighlightedLine = -1
  ClearForm
  ReloadInvoiceModifyForm
  LoadInvoiceLineArray strLineArray
  CountInvoiceLines
  DisplayInvoiceLines 1
  booCheckForChanges = False
  booCheckForLineChanges = False
End Sub

Private Sub cmdDeleteUndelete_Click()
  
  booCheckForLineChanges = True
  
  Dim strAction As String
  Dim strSplits() As String
  strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
  
  If strSplits(11) = "SubItem" Then
    Dim strPrevSplits() As String
    Dim i As Integer
    
    i = intActualHighlightedLine - 1
    Do
      strPrevSplits = Split(strLineArray(i), "<spliter>")
      i = i - 1
    Loop Until strPrevSplits(11) = "Group"
    
    If InStr(1, strPrevSplits(12), "Deleted") > 0 Then
      MsgBox "You can't undelete a subitem in a deleted group item"
      Exit Sub
    End If
  End If
  
  If InStr(1, strSplits(12), "Deleted") > 0 Then
    strLineArray(intActualHighlightedLine) = _
      Left(strLineArray(intActualHighlightedLine), _
           Len(strLineArray(intActualHighlightedLine)) - 7)
    cmdEditLine.Enabled = True
    cmdDeleteUndelete.Caption = "Delete Line"
    strAction = "Undelete"
  Else
    strLineArray(intActualHighlightedLine) = _
      strLineArray(intActualHighlightedLine) & "Deleted"
    cmdEditLine.Enabled = False
    cmdDeleteUndelete.Caption = "Undelete Line"
    strAction = "Delete"
  End If
  
  DisplayLine strLineArray(intActualHighlightedLine), intHighlightedLine
  
  If strSplits(11) = "Group" Then _
    ChangeSubLines intActualHighlightedLine, strAction
End Sub


Private Sub cmdEditLine_Click()
  
  strLastCommand = "EditLine"
  
  SetInvoiceLineInfo strLineArray(intActualHighlightedLine)
  
  If Not booEditInvoiceLineFormLoaded Then Load frmEditInvoiceLine
  frmEditInvoiceLine.Show
End Sub


Private Sub cmdAddLineAfter_Click()
  strLastCommand = "NewLineAfter"
  
  If intNumInvoiceLines = 200 Then
    MsgBox "This program is limited to 200 invoice lines and sublines" & _
      vbCrLf & vbCrLf & "You can't add any more lines"
    Exit Sub
  End If
  
  Dim strSplits() As String
  Dim strItemType As String
  Dim strNextItemType As String
  Dim booNewItem As Boolean
  
  strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
  strItemType = strSplits(11)
  booNewItem = (InStr(1, strSplits(12), "New") > 0)
  
  If intActualHighlightedLine = intNumInvoiceLines Then
    strNextItemType = ""
  Else
    strSplits = Split(strLineArray(intActualHighlightedLine + 1), "<spliter>")
    strNextItemType = strSplits(11)
  End If
  
  If (strItemType = "SubItem" Or _
      (strItemType = "Group" And Not booNewItem)) Then
    If strNextItemType = "SubItem" Then
      If InDeletedGroup(intActualHighlightedLine) Then
        MsgBox "You can't add a new sub line to a group that's been deleted"
        Exit Sub
      Else
        SetInvoiceLineInfo "SubItem,New"
      End If
    Else
      Dim vmbResponse As VbMsgBoxResult
      vmbResponse = vbNo
      If Not InDeletedGroup(intActualHighlightedLine) Then
        vmbResponse = MsgBox(vbCrLf & "Add new line as a group sub line?", vbYesNo)
      End If
      
      If vmbResponse = vbYes Then
        SetInvoiceLineInfo "SubItem,New"
      Else
        SetInvoiceLineInfo "Item,New"
      End If
    End If
  Else
    SetInvoiceLineInfo "Item,New"
  End If
  
  If Not booEditInvoiceLineFormLoaded Then Load frmEditInvoiceLine
  frmEditInvoiceLine.Show
End Sub


Private Sub cmdAddLineBefore_Click()
  strLastCommand = "NewLineBefore"
  
  If intNumInvoiceLines = 200 Then
    MsgBox "This program is limited to 200 invoice lines and sublines" & _
      vbCrLf & vbCrLf & "You can't add any more lines"
    Exit Sub
  End If
  
  If intActualHighlightedLine = 1 Then
    SetInvoiceLineInfo "Item,New"
  Else
    Dim strSplits() As String
    strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
    
    If strSplits(11) = "SubItem" Then
      If InDeletedGroup(intActualHighlightedLine) Then
        MsgBox "You can't add a new sub line to a group that's been deleted"
        Exit Sub
      Else
        SetInvoiceLineInfo "SubItem,New"
      End If
    Else
      strSplits = Split(strLineArray(intActualHighlightedLine - 1), "<spliter>")
      If (strSplits(11) = "SubItem" And _
          Not InDeletedGroup(intActualHighlightedLine - 1)) Or _
         (strSplits(11) = "Group" And InStr(1, strSplits(12), "New") = 0) Then
        
        Dim vmbResponse As VbMsgBoxResult
        vmbResponse = MsgBox(vbCrLf & "Add new line as a group sub line?", vbYesNo)
        If vmbResponse = vbYes Then
          SetInvoiceLineInfo "SubItem,New"
        Else
          SetInvoiceLineInfo "Item,New"
        End If
      Else
        SetInvoiceLineInfo "Item,New"
      End If
    End If 'strSplits(11) = "SubItem"
  End If
  
  If Not booEditInvoiceLineFormLoaded Then Load frmEditInvoiceLine
  frmEditInvoiceLine.Show
End Sub


Private Sub cmdEditMemo_Click()
  booCheckForChanges = True
End Sub


Private Sub cmdFinish_Click()
  strLastCommand = "ReturnToInvoiceSelection"
  frmInvoiceModify.Hide
End Sub


Private Sub chkPending_Click()
  booCheckForChanges = True
End Sub

Private Sub chkToBePrinted_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbARAccount_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbARAccount_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbClass_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbClass_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbCustomer_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbCustomer_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbCustomerMsg_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbCustomerMsg_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbCustTaxCode_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbCustTaxCode_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbItemSalesTax_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbItemSalesTax_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbSalesRep_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbSalesRep_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbShipMethod_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbShipMethod_Change()
  booCheckForChanges = True
End Sub

Private Sub cmbTerms_Click()
  booCheckForChanges = True
End Sub

Private Sub cmbTerms_Change()
  booCheckForChanges = True
End Sub

Private Sub txtAmount_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtBillAddr1_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillAddr2_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillAddr3_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillAddr4_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillCity_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillCountry_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillPostalCode_Change()
  booCheckForChanges = True
End Sub

Private Sub txtBillState_Change()
  booCheckForChanges = True
End Sub

Private Sub txtDesc_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtDueDate_Change()
  booCheckForChanges = True
End Sub

Private Sub txtFOB_Change()
  booCheckForChanges = True
End Sub

Private Sub txtItemFullName_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtItemLineNumber_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtPONumber_Change()
  booCheckForChanges = True
End Sub

Private Sub txtQuantity_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtRate_Click(Index As Integer)
  HighlightLine Index
End Sub

Private Sub txtRefNumber_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipAddr1_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipAddr2_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipAddr3_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipAddr4_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipCity_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipCountry_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipDate_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipPostalCode_Change()
  booCheckForChanges = True
End Sub

Private Sub txtShipState_Change()
  booCheckForChanges = True
End Sub

Private Sub txtSpace1_Click(Index As Integer)
  HighlightLine (Index Mod 10)
End Sub

Private Sub txtTxnDate_Change()
  booCheckForChanges = True
End Sub

Private Sub vscInvoiceLineScroll_Change()
  If vscInvoiceLineScroll.Enabled = False Then Exit Sub
  UnHighlightLine intHighlightedLine
  DisplayInvoiceLines (vscInvoiceLineScroll.Value)
End Sub

Private Sub vscInvoiceLineScroll_GotFocus()
  cmdDoModify.SetFocus
End Sub

Public Sub ClearForm()
  txtTxnID.Text = ""
  txtEditSequence.Text = ""
  txtRefNumber.Text = ""
  txtTxnDate.Text = ""
  chkPending.Value = 0 'Unchecked
  chkToBePrinted.Value = 0 'Unchecked
  cmbCustomer.Text = ""
  cmbClass.Text = ""
  txtBillAddr1.Text = ""
  txtBillAddr2.Text = ""
  txtBillAddr3.Text = ""
  txtBillAddr4.Text = ""
  txtBillCity.Text = ""
  txtBillState.Text = ""
  txtBillPostalCode.Text = ""
  txtBillCountry.Text = ""
  txtShipAddr1.Text = ""
  txtShipAddr2.Text = ""
  txtShipAddr3.Text = ""
  txtShipAddr4.Text = ""
  txtShipCity.Text = ""
  txtShipState.Text = ""
  txtShipPostalCode.Text = ""
  txtShipCountry.Text = ""
  cmbARAccount.Text = ""
  cmbTerms.Text = ""
  txtPONumber.Text = ""
  txtDueDate.Text = ""
  txtShipDate.Text = ""
  txtFOB.Text = ""
  cmbSalesRep.Text = ""
  cmbShipMethod.Text = ""
  cmbItemSalesTax.Text = ""
  cmbCustTaxCode.Text = ""
  cmbCustomerMsg.Text = ""
  txtMemo.Text = ""
  
  Dim i As Integer
  For i = 0 To 9
    txtItemLineNumber(i).Text = ""
    txtQuantity(i).Text = ""
    txtItemFullName(i).Text = ""
    txtDesc(i).Text = ""
    txtRate(i).Text = ""
    txtAmount(i).Text = ""
  Next
  
  For i = 1 To 200
    strLineArray(i) = Empty
  Next i
  
  'Disable the buttons in case we're reactivating the form since a line
  'hasn't yet been chosen
  cmdEditLine.Enabled = False
  cmdAddLineBefore = False
  cmdAddLineAfter.Enabled = False
  cmdDeleteUndelete.Enabled = False
End Sub


Private Sub CountInvoiceLines()

  Dim booDone As Boolean
  Dim i As Integer
  booDone = False
  i = 1
  Do
    If i = 200 Then
      booDone = True
    ElseIf strLineArray(i + 1) = Empty Then
      booDone = True
    Else
      i = i + 1
    End If
  Loop Until booDone
  
  intNumInvoiceLines = i
  
  If intNumInvoiceLines < 11 Then
    vscInvoiceLineScroll.Enabled = False
    vscInvoiceLineScroll.Visible = False
    vscInvoiceLineScroll.Min = 1
    vscInvoiceLineScroll.Max = 1
  Else
    vscInvoiceLineScroll.Enabled = True
    vscInvoiceLineScroll.Visible = True
    vscInvoiceLineScroll.Min = 1
    vscInvoiceLineScroll.Max = intNumInvoiceLines - 9
  End If
End Sub


Private Sub DisplayInvoiceLines(intStartLine As Integer)

  intTopDisplayedLine = intStartLine
  
  Dim i As Integer
  Dim intLastLine As Integer
  
  If intNumInvoiceLines < 10 Then
    intLastLine = intNumInvoiceLines - 1
  Else
    intLastLine = 9
  End If
  
  For i = 0 To intLastLine
    DisplayLine strLineArray(intStartLine + i), i
    If intStartLine + i = intActualHighlightedLine Then
      HighlightLine i
    End If
  Next i
End Sub


Private Sub DisplayLine(strLineInfo As String, _
                        intDisplayLine As Integer)

  Dim strSplits() As String
  
  strSplits = Split(strLineInfo, "<spliter>")
  
  txtItemLineNumber(intDisplayLine).Text = strSplits(0)
  txtQuantity(intDisplayLine).Text = strSplits(1)
  txtItemFullName(intDisplayLine).Text = strSplits(2)
  txtDesc(intDisplayLine).Text = strSplits(3)
  txtRate(intDisplayLine).Text = strSplits(4)
  txtAmount(intDisplayLine).Text = strSplits(5)
  
  If InStr(1, strSplits(12), "Original") > 0 Then
    txtItemLineNumber(intDisplayLine).ForeColor = &H80000008
    txtQuantity(intDisplayLine).ForeColor = &H80000008
    txtItemFullName(intDisplayLine).ForeColor = &H80000008
    txtDesc(intDisplayLine).ForeColor = &H80000008
    txtRate(intDisplayLine).ForeColor = &H80000008
    txtAmount(intDisplayLine).ForeColor = &H80000008
  ElseIf InStr(1, strSplits(12), "New") > 0 Then
    txtItemLineNumber(intDisplayLine).ForeColor = &HC000&
    txtQuantity(intDisplayLine).ForeColor = &HC000&
    txtItemFullName(intDisplayLine).ForeColor = &HC000&
    txtDesc(intDisplayLine).ForeColor = &HC000&
    txtRate(intDisplayLine).ForeColor = &HC000&
    txtAmount(intDisplayLine).ForeColor = &HC000&
  ElseIf InStr(1, strSplits(12), "Changed") > 0 Then
    txtItemLineNumber(intDisplayLine).ForeColor = &HFF&
    txtQuantity(intDisplayLine).ForeColor = &HFF&
    txtItemFullName(intDisplayLine).ForeColor = &HFF&
    txtDesc(intDisplayLine).ForeColor = &HFF&
    txtRate(intDisplayLine).ForeColor = &HFF&
    txtAmount(intDisplayLine).ForeColor = &HFF&
  End If
  
  If InStr(1, strSplits(12), "Deleted") > 0 Then
    txtItemLineNumber(intDisplayLine).FontStrikethru = True
    txtQuantity(intDisplayLine).FontStrikethru = True
    txtItemFullName(intDisplayLine).FontStrikethru = True
    txtDesc(intDisplayLine).FontStrikethru = True
    txtRate(intDisplayLine).FontStrikethru = True
    txtAmount(intDisplayLine).FontStrikethru = True
  Else
    txtItemLineNumber(intDisplayLine).FontStrikethru = False
    txtQuantity(intDisplayLine).FontStrikethru = False
    txtItemFullName(intDisplayLine).FontStrikethru = False
    txtDesc(intDisplayLine).FontStrikethru = False
    txtRate(intDisplayLine).FontStrikethru = False
    txtAmount(intDisplayLine).FontStrikethru = False
  End If

  If strSplits(11) = "SubItem" Then
    txtItemLineNumber(intDisplayLine).FontBold = False
  Else
    txtItemLineNumber(intDisplayLine).FontBold = True
  End If

  If strSplits(11) = "Group" Then
    txtItemFullName(intDisplayLine).FontBold = True
    txtDesc(intDisplayLine).FontBold = True
  Else
    txtItemFullName(intDisplayLine).FontBold = False
    txtDesc(intDisplayLine).FontBold = False
  End If
End Sub


Private Sub HighlightLine(intDisplayLine As Integer)
  If intDisplayLine >= intNumInvoiceLines Then Exit Sub
  
  If intHighlightedLine > -1 Then UnHighlightLine intHighlightedLine
  
  txtItemLineNumber(intDisplayLine).BackColor = &HFFFF00
  txtQuantity(intDisplayLine).BackColor = &HFFFF00
  txtItemFullName(intDisplayLine).BackColor = &HFFFF00
  txtDesc(intDisplayLine).BackColor = &HFFFF00
  txtRate(intDisplayLine).BackColor = &HFFFF00
  txtAmount(intDisplayLine).BackColor = &HFFFF00
  txtSpace1(intDisplayLine).BackColor = &HFFFF00
  txtSpace1(intDisplayLine + 10).BackColor = &HFFFF00
  txtSpace1(intDisplayLine + 20).BackColor = &HFFFF00
  
  intActualHighlightedLine = vscInvoiceLineScroll.Value + intDisplayLine
  intHighlightedLine = intDisplayLine
  
  Dim strSplits() As String
  strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
  
  cmdAddLineBefore.Enabled = True
  cmdAddLineAfter.Enabled = True
  cmdDeleteUndelete.Enabled = True
  
  If InStr(1, strSplits(12), "Deleted") > 0 Then
    cmdDeleteUndelete.Caption = "Undelete Line"
    cmdEditLine.Enabled = False
  Else
    cmdDeleteUndelete.Caption = "Delete Line"
    cmdEditLine.Enabled = True
  End If
End Sub


Private Sub UnHighlightLine(intDisplayLine As Integer)
  If intHighlightedLine > -1 Then
    txtItemLineNumber(intDisplayLine).BackColor = &H80000005
    txtQuantity(intDisplayLine).BackColor = &H80000005
    txtItemFullName(intDisplayLine).BackColor = &H80000005
    txtDesc(intDisplayLine).BackColor = &H80000005
    txtRate(intDisplayLine).BackColor = &H80000005
    txtAmount(intDisplayLine).BackColor = &H80000005
    txtSpace1(intDisplayLine).BackColor = &H80000005
    txtSpace1(intDisplayLine + 10).BackColor = &H80000005
    txtSpace1(intDisplayLine + 20).BackColor = &H80000005
    
    cmdAddLineBefore.Enabled = False
    cmdAddLineAfter.Enabled = False
    cmdDeleteUndelete.Enabled = False
    cmdEditLine.Enabled = False
  
    intHighlightedLine = -1
  End If
End Sub


Private Sub InsertInvoiceLine(strInvoiceLine As String, _
                              intLineNumber As Integer)

  If intLineNumber > intNumInvoiceLines Then
    strLineArray(intNumInvoiceLines + 1) = GetInvoiceLineInfo
  Else
    Dim i As Integer
    For i = intNumInvoiceLines + 1 To intLineNumber + 1 Step -1
      strLineArray(i) = strLineArray(i - 1)
    Next i
    strLineArray(intLineNumber) = strInvoiceLine
  End If
  
  intNumInvoiceLines = intNumInvoiceLines + 1
  If intNumInvoiceLines = 11 Then
    vscInvoiceLineScroll.Max = 2
    vscInvoiceLineScroll.Enabled = True
    vscInvoiceLineScroll.Visible = True
    vscInvoiceLineScroll.Value = intTopDisplayedLine
  ElseIf intNumInvoiceLines > 11 Then
    vscInvoiceLineScroll.Max = vscInvoiceLineScroll.Max + 1
  End If
End Sub


Private Sub SaveOriginalValues()
  strTxnID = txtTxnID.Text
  strEditSequence = txtEditSequence.Text
  strRefNumber = txtRefNumber.Text
  strTxnDate = txtTxnDate.Text
  intPending = chkPending.Value
  intToBePrinted = chkToBePrinted.Value
  strCustomer = cmbCustomer.Text
  strClass = cmbClass.Text
  strBillAddr1 = txtBillAddr1.Text
  strBillAddr2 = txtBillAddr2.Text
  strBillAddr3 = txtBillAddr3.Text
  strBillAddr4 = txtBillAddr4.Text
  strBillCity = txtBillCity.Text
  strBillState = txtBillState.Text
  strBillPostalCode = txtBillPostalCode.Text
  strBillCountry = txtBillCountry.Text
  strShipAddr1 = txtShipAddr1.Text
  strShipAddr2 = txtShipAddr2.Text
  strShipAddr3 = txtShipAddr3.Text
  strShipAddr4 = txtShipAddr4.Text
  strShipCity = txtShipCity.Text
  strShipState = txtShipState.Text
  strShipPostalCode = txtShipPostalCode.Text
  strShipCountry = txtShipCountry.Text
  strARAccount = cmbARAccount.Text
  strTerms = cmbTerms.Text
  strPONumber = txtPONumber.Text
  strDueDate = txtDueDate.Text
  strShipDate = txtShipDate.Text
  strFOB = txtFOB.Text
  strSalesRep = cmbSalesRep.Text
  strShipMethod = cmbShipMethod.Text
  strItemSalesTax = cmbItemSalesTax.Text
  strCustTaxCode = cmbCustTaxCode.Text
  strCustomerMsg = cmbCustomerMsg.Text
  strMemo = txtMemo.Text
End Sub


Private Sub ReloadInvoiceModifyForm()
  txtTxnID.Text = strTxnID
  txtEditSequence.Text = strEditSequence
  txtRefNumber.Text = strRefNumber
  txtTxnDate.Text = strTxnDate
  chkPending.Value = intPending
  chkToBePrinted.Value = intToBePrinted
  cmbCustomer.Text = strCustomer
  cmbClass.Text = strClass
  txtBillAddr1.Text = strBillAddr1
  txtBillAddr2.Text = strBillAddr2
  txtBillAddr3.Text = strBillAddr3
  txtBillAddr4.Text = strBillAddr4
  txtBillCity.Text = strBillCity
  txtBillState.Text = strBillState
  txtBillPostalCode.Text = strBillPostalCode
  txtBillCountry.Text = strBillCountry
  txtShipAddr1.Text = strShipAddr1
  txtShipAddr2.Text = strShipAddr2
  txtShipAddr3.Text = strShipAddr3
  txtShipAddr4.Text = strShipAddr4
  txtShipCity.Text = strShipCity
  txtShipState.Text = strShipState
  txtShipPostalCode.Text = strShipPostalCode
  txtShipCountry.Text = strShipCountry
  cmbARAccount.Text = strARAccount
  cmbTerms.Text = strTerms
  txtPONumber.Text = strPONumber
  txtDueDate.Text = strDueDate
  txtShipDate.Text = strShipDate
  txtFOB.Text = strFOB
  cmbSalesRep.Text = strSalesRep
  cmbShipMethod.Text = strShipMethod
  cmbItemSalesTax.Text = strItemSalesTax
  cmbCustTaxCode.Text = strCustTaxCode
  cmbCustomerMsg.Text = strCustomerMsg
  txtMemo.Text = strMemo
End Sub


Private Function InvoiceChanged() As Boolean

  booBillAddressChanged = False
  booShipAddressChanged = False
  booInvoiceBodyChanged = False
  booInvoiceLinesChanged = False
  
  If Not (booCheckForChanges Or booCheckForLineChanges) Then
    InvoiceChanged = False
    Exit Function
  End If
  Dim booInvoiceChanged As Boolean
  booInvoiceChanged = False
  
  If Trim(strRefNumber) <> Trim(txtRefNumber.Text) Or _
     Trim(strTxnDate) <> Trim(txtTxnDate.Text) Or _
     intPending <> chkPending.Value Or _
     intToBePrinted <> chkToBePrinted.Value Or _
     Trim(strCustomer) <> Trim(cmbCustomer.Text) Or _
     Trim(strClass) <> Trim(cmbClass.Text) Or _
     Trim(strARAccount) <> Trim(cmbARAccount.Text) Or _
     Trim(strTerms) <> Trim(cmbTerms.Text) Or _
     Trim(strPONumber) <> Trim(txtPONumber.Text) Or _
     Trim(strDueDate) <> Trim(txtDueDate.Text) Or _
     Trim(strShipDate) <> Trim(txtShipDate.Text) Or _
     Trim(strFOB) <> Trim(txtFOB.Text) Or _
     Trim(strSalesRep) <> Trim(cmbSalesRep.Text) Or _
     Trim(strShipMethod) <> Trim(cmbShipMethod.Text) Or _
     Trim(strItemSalesTax) <> Trim(cmbItemSalesTax.Text) Or _
     Trim(strCustTaxCode) <> Trim(cmbCustTaxCode.Text) Or _
     Trim(strCustomerMsg) <> Trim(cmbCustomerMsg.Text) Or _
     Trim(strMemo) <> Trim(txtMemo.Text) Then
    booInvoiceChanged = True
    booInvoiceBodyChanged = True
  End If

  If Trim(strBillAddr1) <> Trim(txtBillAddr1.Text) Or _
     Trim(strBillAddr2) <> Trim(txtBillAddr2.Text) Or _
     Trim(strBillAddr3) <> Trim(txtBillAddr3.Text) Or _
     Trim(strBillAddr4) <> Trim(txtBillAddr4.Text) Or _
     Trim(strBillCity) <> Trim(txtBillCity.Text) Or _
     Trim(strBillState) <> Trim(txtBillState.Text) Or _
     Trim(strBillPostalCode) <> Trim(txtBillPostalCode.Text) Or _
     Trim(strBillCountry) <> Trim(txtBillCountry.Text) Then
    booBillAddressChanged = True
    booInvoiceChanged = True
    booInvoiceBodyChanged = True
  End If

  If Trim(strShipAddr1) <> Trim(txtShipAddr1.Text) Or _
     Trim(strShipAddr2) <> Trim(txtShipAddr2.Text) Or _
     Trim(strShipAddr3) <> Trim(txtShipAddr3.Text) Or _
     Trim(strShipAddr4) <> Trim(txtShipAddr4.Text) Or _
     Trim(strShipCity) <> Trim(txtShipCity.Text) Or _
     Trim(strShipState) <> Trim(txtShipState.Text) Or _
     Trim(strShipPostalCode) <> Trim(txtShipPostalCode.Text) Or _
     Trim(strShipCountry) <> Trim(txtShipCountry.Text) Then
    booShipAddressChanged = True
    booInvoiceChanged = True
    booInvoiceBodyChanged = True
  End If
  
  Dim strSplits() As String
  Dim i As Integer
  For i = 1 To intNumInvoiceLines
    strSplits = Split(strLineArray(i), "<spliter>")
    If strSplits(12) = "New" Or strSplits(12) = "Changed" Or _
       strSplits(12) = "OriginalDeleted" Then
      booInvoiceLinesChanged = True
    End If
  Next i
  InvoiceChanged = booInvoiceChanged Or booInvoiceLinesChanged
End Function


Private Function InvoiceChangeString() As String

  Dim strInvoiceChangeString As String
  
  strInvoiceChangeString = strInvoiceChangeString & _
    "<TxnID>" & txtTxnID.Text & "</TxnID>" & _
    "<EditSequence>" & Trim(txtEditSequence.Text) & "</EditSequence>"
  
  If booInvoiceBodyChanged Then
    If cmbCustomer.Text <> strCustomer Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<CustomerRef><FullName>" & Trim(cmbCustomer.Text) & _
        "</FullName></CustomerRef>"
    End If

    If cmbClass.Text <> strClass Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ClassRef><FullName>" & Trim(cmbClass.Text) & "</FullName></ClassRef>"
    End If

    If cmbARAccount.Text <> strARAccount Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ARAccountRef><FullName>" & Trim(cmbARAccount.Text) & "</FullName></ARAccountRef>"
    End If

    If txtTxnDate.Text <> strTxnDate Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<TxnDate>" & Trim(txtTxnDate.Text) & "</TxnDate>"
    End If

    If txtRefNumber.Text <> strRefNumber Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<RefNumberCaseSensitive>" & Trim(txtRefNumber.Text) & "</RefNumberCaseSensitive>"
    End If

    If booBillAddressChanged Then
      strInvoiceChangeString = strInvoiceChangeString & "<BillAddress>"
      If txtBillAddr1.Text <> strBillAddr1 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr1>" & Trim(txtBillAddr1.Text) & "</Addr1>"
      End If

      If txtBillAddr2.Text <> strBillAddr2 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr2>" & Trim(txtBillAddr2.Text) & "</Addr2>"
      End If

      If txtBillAddr3.Text <> strBillAddr3 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr3>" & Trim(txtBillAddr3.Text) & "</Addr3>"
      End If

      If txtBillAddr4.Text <> strBillAddr4 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr4>" & Trim(txtBillAddr4.Text) & "</Addr4>"
      End If

      If txtBillCity.Text <> strBillCity Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<City>" & Trim(txtBillCity.Text) & "</City>"
      End If

      If txtBillState.Text <> strBillState Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<State>" & Trim(txtBillState.Text) & "</State>"
      End If

      If txtBillPostalCode.Text <> strBillPostalCode Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<PostalCode>" & Trim(txtBillPostalCode.Text) & "</PostalCode>"
      End If

      If txtBillCountry.Text <> strBillCountry Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Country>" & Trim(txtBillCountry.Text) & "</Country>"
      End If

      strInvoiceChangeString = strInvoiceChangeString & "</BillAddress>"
    End If ' If booBillAddressChanged

    If booShipAddressChanged Then
      strInvoiceChangeString = strInvoiceChangeString & "<ShipAddress>"
      If txtShipAddr1.Text <> strShipAddr1 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr1>" & Trim(txtShipAddr1.Text) & "</Addr1>"
      End If

      If txtShipAddr2.Text <> strShipAddr2 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr2>" & Trim(txtShipAddr2.Text) & "</Addr2>"
      End If

      If txtShipAddr3.Text <> strShipAddr3 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr3>" & Trim(txtShipAddr3.Text) & "</Addr3>"
      End If

      If txtShipAddr4.Text <> strShipAddr4 Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Addr4>" & Trim(txtShipAddr4.Text) & "</Addr4>"
      End If

      If txtShipCity.Text <> strShipCity Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<City>" & Trim(txtShipCity.Text) & "</City>"
      End If

      If txtShipState.Text <> strShipState Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<State>" & Trim(txtShipState.Text) & "</State>"
      End If

      If txtShipPostalCode.Text <> strShipPostalCode Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<PostalCode>" & Trim(txtShipPostalCode.Text) & "</PostalCode>"
      End If

      If txtShipCountry.Text <> strShipCountry Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<Country>" & Trim(txtShipCountry.Text) & "</Country>"
      End If

      strInvoiceChangeString = strInvoiceChangeString & "</ShipAddress>"
    End If ' If booShipAddressChanged

    If chkPending.Value <> intPending Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<IsPending>" & chkPending.Value & "</IsPending>"
    End If

    If txtPONumber.Text <> strPONumber Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<PONumber>" & Trim(txtPONumber.Text) & "</PONumber>"
    End If

    If cmbTerms.Text <> strTerms Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<TermsRef><FullName>" & Trim(cmbTerms.Text) & "</FullName></TermsRef>"
    End If

    If txtDueDate.Text <> strDueDate Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<DueDate>" & Trim(txtDueDate.Text) & "</DueDate>"
    End If

    If cmbSalesRep.Text <> strSalesRep Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<SalesRepRef><FullName>" & Trim(cmbSalesRep.Text) & "</FullName></SalesRepRef>"
    End If

    If txtFOB.Text <> strFOB Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<FOB>" & Trim(txtFOB.Text) & "</FOB>"
    End If

    If txtShipDate.Text <> strShipDate Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ShipDate>" & Trim(txtShipDate.Text) & "</ShipDate>"
    End If

    If cmbShipMethod.Text <> strShipMethod Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ShipMethodRef><FullName>" & Trim(cmbShipMethod.Text) & "</FullName></ShipMethodRef>"
    End If

    If cmbItemSalesTax.Text <> strItemSalesTax Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ItemSalesTaxRef><FullName>" & Trim(cmbItemSalesTax.Text) & "</FullName></ItemSalesTaxRef>"
    End If

    If txtMemo.Text <> strMemo Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<Memo>" & Trim(txtMemo.Text) & "</Memo>"
    End If

    If cmbCustomerMsg.Text <> strCustomerMsg Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<CustomerMsgRef><FullName>" & Trim(cmbCustomerMsg.Text) & "</FullName></CustomerMsgRef>"
    End If

    If chkToBePrinted.Value <> intToBePrinted Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<IsToBePrinted>" & chkToBePrinted.Value & "</IsToBePrinted>"
    End If

    If cmbCustTaxCode.Text <> strCustTaxCode Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<CustomerSalesTaxCodeRef><FullName>" & Trim(cmbCustTaxCode.Text) & "</FullName></CustomerSalesTaxCodeRef>"
    End If
  End If
  
  If booInvoiceLinesChanged Then
    Dim strSplits() As String
    Dim strGroupLineSplits() As String
    Dim booDone As Boolean
    Dim booLineLoopDone As Boolean
    Dim booIncludeGroupLines As Boolean
    Dim i As Integer
    Dim j As Integer
    booDone = False
    i = 1
    Do
      strSplits = Split(strLineArray(i), "<spliter>")
      If strSplits(11) = "Group" Then
        If i = 200 And InStr(1, strSplits(12), "Deleted") <= 0 Then
          AddInvoiceLineInfo strInvoiceChangeString, i, False
          InvoiceChangeString = strInvoiceChangeString
          Exit Function
        End If
        
        j = i + 1
        booLineLoopDone = False
        booIncludeGroupLines = False
        Do While Not booLineLoopDone
          strGroupLineSplits = Split(strLineArray(j), "<spliter>")
          If strGroupLineSplits(11) <> "SubItem" Then
            booLineLoopDone = True
          Else
            If strGroupLineSplits(12) <> "NewDeleted" And _
               strGroupLineSplits(12) <> "Original" Then
              booIncludeGroupLines = True
            End If
            
            If j = UBound(strLineArray) Then
              booLineLoopDone = True
            ElseIf strLineArray(j + 1) = Empty Then
              'set the value of j to j + 1 anyway to avoid having the last
              'item in a group be repeated if it's the last item on an invoice
              j = j + 1
              booLineLoopDone = True
            Else
              j = j + 1
            End If
          End If
        Loop
        
        If InStr(1, strSplits(12), "Deleted") <= 0 Then
          AddInvoiceLineInfo strInvoiceChangeString, i, booIncludeGroupLines
        End If
        
        i = j - 1
      Else ' for If strSplits(11) = "Group"
        If strSplits(12) = "Original" Then
          strInvoiceChangeString = strInvoiceChangeString & _
            "<InvoiceLineMod><TxnLineID>" & strSplits(0) & _
            "</TxnLineID></InvoiceLineMod>"
        ElseIf strSplits(12) = "Changed" Or strSplits(12) = "New" Then
          AddInvoiceLineInfo strInvoiceChangeString, i, False
        End If ' for If strSplits(12) = "Original"
      End If ' for If strSplits(11) = "Group"
      
      If i = 200 Then
        booDone = True
      Else
        If strLineArray(i + 1) = Empty Then
          booDone = True
        Else
          i = i + 1
        End If
      End If
    Loop Until booDone
  End If
  
  InvoiceChangeString = strInvoiceChangeString
End Function


Private Sub AddInvoiceLineInfo(strInvoiceChangeString As String, _
                               intInvoiceLineNum As Integer, _
                               booAddGroupLines As Boolean)

  Dim booGroupLine As Boolean
  booGroupLine = False
  Dim strSplits() As String
  strSplits = Split(strLineArray(intInvoiceLineNum), "<spliter>")
  
  If strSplits(11) = "Group" Then
    strInvoiceChangeString = strInvoiceChangeString & _
      "<InvoiceLineGroupMod><TxnLineID>"
    booGroupLine = True
  Else
    If InStr(1, strSplits(2), " - Group Item") Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<InvoiceLineGroupMod><TxnLineID>"
      booGroupLine = True
    Else
      strInvoiceChangeString = strInvoiceChangeString & _
        "<InvoiceLineMod><TxnLineID>"
    End If
  End If
  
  If Left(strSplits(0), 2) = "-1" Then
    strInvoiceChangeString = strInvoiceChangeString & "-1</TxnLineID>"
  Else
    strInvoiceChangeString = strInvoiceChangeString & _
    strSplits(0) & "</TxnLineID>"
  End If

  If Trim(strSplits(2)) <> Empty Then
    If InStr(1, strSplits(2), " - Group Item") > 0 Or _
       strSplits(11) = "Group" Then
      If strSplits(12) <> "Original" Then
        strInvoiceChangeString = strInvoiceChangeString & _
          "<ItemGroupRef><FullName>" & _
          Replace(Trim(strSplits(2)), " - Group Item", "") & _
          "</FullName></ItemGroupRef>"
      End If
    Else
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ItemRef><FullName>" & Trim(strSplits(2)) & "</FullName></ItemRef>"
    End If
  End If
  
  If Trim(strSplits(3)) <> Empty And strSplits(11) <> "Group" And _
     InStr(1, strSplits(2), " - Group Item") = 0 Then
    strInvoiceChangeString = strInvoiceChangeString & _
    "<Desc>" & Trim(strSplits(3)) & "</Desc>"
  End If
  
  If Trim(strSplits(1)) <> Empty And strSplits(12) <> "Original" Then
    strInvoiceChangeString = strInvoiceChangeString & _
    "<Quantity>" & Trim(strSplits(1)) & "</Quantity>"
  End If
  
  If Trim(strSplits(4)) <> Empty And strSplits(11) <> "Group" Then
    If strSplits(9) = "RatePercent" Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<RatePercent>" & Trim(strSplits(4)) & "</RatePercent>"
    Else
      strInvoiceChangeString = strInvoiceChangeString & _
        "<Rate>" & Trim(strSplits(4)) & "</Rate>"
    End If
  End If
  
  If Trim(strSplits(6)) <> Empty And strSplits(11) <> "Group" Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<ClassRef><FullName>" & Trim(strSplits(6)) & "</FullName></ClassRef>"
  End If
  
  If Trim(strSplits(5)) <> Empty And strSplits(11) <> "Group" Then
    strInvoiceChangeString = strInvoiceChangeString & _
    "<Amount>" & Trim(strSplits(5)) & "</Amount>"
  End If
  
  If Trim(strSplits(7)) <> Empty And strSplits(11) <> "Group" Then
    strInvoiceChangeString = strInvoiceChangeString & _
    "<ServiceDate>" & Trim(strSplits(7)) & "</ServiceDate>"
  End If
  
  If Trim(strSplits(8)) <> Empty And strSplits(11) <> "Group" Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<SalesTaxCodeRef><FullName>" & Trim(strSplits(8)) & _
        "</FullName></SalesTaxCodeRef>"
  End If
  
  If Trim(strSplits(10)) <> Empty And strSplits(11) <> "Group" Then
      strInvoiceChangeString = strInvoiceChangeString & _
        "<OverrideItemAccountRef><FullName>" & Trim(strSplits(10)) & _
        "</FullName></OverrideItemAccountRef>"
  End If
  
  If booAddGroupLines Then
    Dim j As Integer
    Dim booProcessGroupLines As Boolean
    Dim strGroupLineSplits() As String
    
    If intInvoiceLineNum = 200 Then
      booProcessGroupLines = False
    ElseIf InStr(1, strLineArray(intInvoiceLineNum + 1), "SubItem") <= 0 Then
      booProcessGroupLines = False
    Else
      booProcessGroupLines = True
      j = intInvoiceLineNum + 1
    End If
    
    Do While booProcessGroupLines
      strGroupLineSplits = Split(strLineArray(j), "<spliter>")
      If strGroupLineSplits(11) <> "SubItem" Then
        booProcessGroupLines = False
      Else
        If strGroupLineSplits(12) = "Original" Then
          strInvoiceChangeString = strInvoiceChangeString & _
            "<InvoiceLineMod><TxnLineID>" & strGroupLineSplits(0) & _
            "</TxnLineID></InvoiceLineMod>"
        ElseIf strGroupLineSplits(12) = "New" Or _
               strGroupLineSplits(12) = "Changed" Then
          AddInvoiceLineInfo strInvoiceChangeString, j, False
        End If

        If j = UBound(strLineArray) Then
          booProcessGroupLines = False
        ElseIf strLineArray(j + 1) = Empty Then
          booProcessGroupLines = False
        Else
          j = j + 1
        End If
      End If
    Loop
  End If
  
  If booGroupLine Then
    strInvoiceChangeString = strInvoiceChangeString & "</InvoiceLineGroupMod>"
  Else
    strInvoiceChangeString = strInvoiceChangeString & "</InvoiceLineMod>"
  End If
End Sub


Private Sub ChangeSubLines(intGroupLine As Integer, _
                           strAction As String)

  If intGroupLine = UBound(strLineArray) Then Exit Sub
  
  Dim i As Integer
  Dim booDone As Boolean
  Dim strSplits() As String
  
  i = intGroupLine + 1
  booDone = False
  Do
    strSplits = Split(strLineArray(i), "<spliter>")
    If strSplits(11) <> "SubItem" Then
      booDone = True
    Else
      If strAction = "Delete" Then
        strLineArray(i) = strLineArray(i) & "Deleted"
      Else
        strLineArray(i) = Left(strLineArray(i), Len(strLineArray(i)) - 7)
      End If
    
      If intHighlightedLine + (i - intActualHighlightedLine) < 10 Then
        DisplayLine _
          strLineArray(i), intHighlightedLine + (i - intActualHighlightedLine)
      End If
      
      If i = UBound(strLineArray) Then
        booDone = True
      ElseIf strLineArray(i + 1) = Empty Then
        booDone = True
      Else
        i = i + 1
      End If
    End If
  Loop Until booDone
End Sub


Private Function InDeletedGroup(intArrayLine As Integer) As Boolean

  Dim strSplits() As String
  Dim i As Integer
  
  strSplits = Split(strLineArray(intArrayLine), "<spliter>")
  If strSplits(11) <> "SubItem" Then
    InDeletedGroup = False
    Exit Function
  End If
  
  i = intArrayLine
  Do
    i = i - 1
    strSplits = Split(strLineArray(i), "<spliter>")
  Loop Until strSplits(11) = "Group"
  
  InDeletedGroup = (InStr(1, strSplits(12), "Deleted") > 0)
End Function


VERSION 5.00
Begin VB.Form frmPurchaseOrderDetail 
   Caption         =   "Purchase Order Details"
   ClientHeight    =   7395
   ClientLeft      =   3360
   ClientTop       =   2355
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9270
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   1095
      Left            =   7320
      TabIndex        =   134
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to open PO list"
      Height          =   1095
      Left            =   5160
      TabIndex        =   133
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCloseLine 
      Caption         =   "Close highlighted line without receiving remaining quantity"
      Height          =   495
      Left            =   120
      TabIndex        =   132
      Top             =   6720
      Width           =   4815
   End
   Begin VB.CommandButton cmdCloseAll 
      Caption         =   "Close entire purchase order without receiving open items"
      Height          =   495
      Left            =   120
      TabIndex        =   131
      Top             =   6120
      Width           =   4815
   End
   Begin VB.CommandButton cmdReceiveOne 
      Caption         =   "Set quantity orderd to quantity received and bill for the difference for highlighted line"
      Height          =   495
      Left            =   120
      TabIndex        =   130
      Top             =   5520
      Width           =   4815
   End
   Begin VB.CommandButton cmdReceiveAll 
      Caption         =   "Set quantities orderd to quantities received and bill for the difference for entire purchase order"
      Height          =   495
      Left            =   120
      TabIndex        =   129
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   " Purchase Order Lines "
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   8655
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   121
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   7080
         TabIndex        =   120
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   6360
         TabIndex        =   119
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   6240
         TabIndex        =   118
         Top             =   2760
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   5280
         TabIndex        =   116
         Top             =   2760
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   4440
         TabIndex        =   114
         Top             =   2760
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   113
         Top             =   2760
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   9
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   2760
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   110
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   7080
         TabIndex        =   109
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   6360
         TabIndex        =   108
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   6240
         TabIndex        =   107
         Top             =   2520
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   5280
         TabIndex        =   105
         Top             =   2520
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   4440
         TabIndex        =   103
         Top             =   2520
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   102
         Top             =   2520
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   8
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   2520
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   99
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   7080
         TabIndex        =   98
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   6360
         TabIndex        =   97
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   6240
         TabIndex        =   96
         Top             =   2280
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   5280
         TabIndex        =   94
         Top             =   2280
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   4440
         TabIndex        =   92
         Top             =   2280
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   91
         Top             =   2280
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   7
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   2280
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   88
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   7080
         TabIndex        =   87
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   6360
         TabIndex        =   86
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   6240
         TabIndex        =   85
         Top             =   2040
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   5280
         TabIndex        =   83
         Top             =   2040
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   4440
         TabIndex        =   81
         Top             =   2040
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   80
         Top             =   2040
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   2040
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   77
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   7080
         TabIndex        =   76
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   6360
         TabIndex        =   75
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   6240
         TabIndex        =   74
         Top             =   1800
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   5280
         TabIndex        =   72
         Top             =   1800
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   4440
         TabIndex        =   70
         Top             =   1800
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   69
         Top             =   1800
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   5
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1800
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   66
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   7080
         TabIndex        =   65
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   6360
         TabIndex        =   64
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   6240
         TabIndex        =   63
         Top             =   1560
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   61
         Top             =   1560
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   4440
         TabIndex        =   59
         Top             =   1560
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   58
         Top             =   1560
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1560
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   7080
         TabIndex        =   54
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   6360
         TabIndex        =   53
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   6240
         TabIndex        =   52
         Top             =   1320
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   50
         Top             =   1320
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   4440
         TabIndex        =   48
         Top             =   1320
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   47
         Top             =   1320
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1320
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   6240
         TabIndex        =   41
         Top             =   1080
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   39
         Top             =   1080
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   4440
         TabIndex        =   37
         Top             =   1080
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   36
         Top             =   1080
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   7080
         TabIndex        =   32
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   30
         Top             =   840
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   28
         Top             =   840
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   26
         Top             =   840
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   25
         Top             =   840
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   1435
      End
      Begin VB.TextBox txtClosed 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   7920
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   7080
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSpace4 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   19
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtSpace3 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   17
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSpace2 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox txtSpace1 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   130
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   280
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1435
      End
      Begin VB.Label Label6 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Description"
         Height          =   255
         Left            =   1680
         TabIndex        =   127
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   4680
         TabIndex        =   126
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5880
         TabIndex        =   125
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Amount"
         Height          =   255
         Left            =   6480
         TabIndex        =   124
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Received"
         Height          =   255
         Left            =   7200
         TabIndex        =   123
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Closed"
         Height          =   255
         Left            =   7920
         TabIndex        =   122
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.VScrollBar vscPOLineScroll 
      Height          =   2415
      LargeChange     =   10
      Left            =   8880
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtVendor 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtRefNumber 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtTxnDate 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   1335
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
   Begin VB.TextBox txtEditSequence 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "PO #"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Date"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "TxnID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Edit Sequence"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPurchaseOrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Dim strPOLines(30) As String
Dim strTxnID As String
Dim strEditSequence As String
Dim strRefNumber As String
Dim strTxnDate As String
Dim strVendor As String

Dim intPOLines As Integer
Dim intHighlightedPOLineRow As Integer
Dim intActualHighlightedPOLine As Integer
Dim intTopDisplayedPOLine As Integer


Private Sub Form_Load()
  ClearstrPOLines
  GetPOInfo strTxnID, strEditSequence, strRefNumber, strTxnDate, _
            strVendor, strPOLines
  
  If strPOLines(1) = Empty Then
    MsgBox "Error encountered getting PO lines, exiting"
    End
  End If
  
  FillPODisplay
  intHighlightedPOLineRow = -1
  intActualHighlightedPOLine = -1
End Sub


Private Sub cmdReturn_Click()
  frmPurchaseOrderSelect.RefreshPOList False
  Unload frmPurchaseOrderDetail
End Sub


Private Sub cmdExit_Click()
  EndSessionCloseConnection
  End
End Sub


Private Sub cmdReceiveAll_Click()
  SetQuantitiesAndBillForRemainingItems _
    strTxnID, strEditSequence, strVendor, strRefNumber, _
    strTxnDate, strPOLines, 0
  frmPurchaseOrderSelect.RefreshPOList False
  Unload frmPurchaseOrderDetail
End Sub


Private Sub cmdReceiveOne_Click()

  If intActualHighlightedPOLine = -1 Then
    MsgBox "You must select a PO line before using this button"
    Exit Sub
  End If

  If LineType(strPOLines(intActualHighlightedPOLine)) = "GroupItem" Then
    MsgBox "You cannot receive a Group Item line" & vbCrLf & vbCrLf & _
      "You can only receive the Item lines in a Group Item"
    Exit Sub
  End If

  SetQuantitiesAndBillForRemainingItems _
    strTxnID, strEditSequence, strVendor, strRefNumber, _
    strTxnDate, strPOLines, intActualHighlightedPOLine
  frmPurchaseOrderSelect.RefreshPOList False
  Unload frmPurchaseOrderDetail
End Sub


Private Sub cmdCloseLine_Click()

  If intActualHighlightedPOLine = -1 Then
    MsgBox "You must select a PO line before using this button"
    Exit Sub
  End If
  
  If LineType(strPOLines(intActualHighlightedPOLine)) = "GroupItem" Then
    MsgBox "You cannot close a Group Item line" & vbCrLf & vbCrLf & _
      "You can only close the Item lines in a Group Item"
    Exit Sub
  End If
  
  If POLineClosed(strPOLines(intActualHighlightedPOLine)) Then
    MsgBox "You cannot close a line that's already closed"
    Exit Sub
  End If
  
  ChangePOLine "Close", strTxnID, strEditSequence, strPOLines, intActualHighlightedPOLine
  UnHighlightPOLineRow intHighlightedPOLineRow
  frmPurchaseOrderSelect.RefreshPOList False
  Unload frmPurchaseOrderDetail
End Sub


Private Sub cmdCloseAll_Click()
  ClosePO txtTxnID.Text, txtEditSequence.Text
  frmPurchaseOrderSelect.RefreshPOList False
  Unload frmPurchaseOrderDetail
End Sub


Private Sub ClearstrPOLines()

  Dim i As Integer
  For i = 1 To UBound(strPOLines)
    strPOLines(i) = Empty
  Next i
End Sub


Private Sub FillPODisplay()

  txtTxnID.Text = strTxnID
  txtEditSequence.Text = strEditSequence
  txtRefNumber.Text = strRefNumber
  txtTxnDate.Text = strTxnDate
  txtVendor.Text = strVendor
  
  Dim strSplits() As String
  Dim i As Integer
  Dim booDone As Boolean
  i = 1
  booDone = False
  
  intTopDisplayedPOLine = 1
  Do
    If i < 11 Then
      strSplits = Split(strPOLines(i), "<spliter>")
      txtItem(i - 1).Text = strSplits(0)
      txtDescription(i - 1).Text = strSplits(1)
      txtQuantity(i - 1).Text = strSplits(2)
      txtRate(i - 1).Text = strSplits(3)
      txtAmount(i - 1).Text = strSplits(5)
      txtReceived(i - 1).Text = strSplits(6)
      txtClosed(i - 1).Text = strSplits(7)
    End If
    
    If i = 30 Then
      booDone = True
    ElseIf strPOLines(i + 1) = Empty Then
      booDone = True
    Else
      i = i + 1
    End If
  Loop Until booDone
  
  intPOLines = i
  
  If intPOLines < 11 Then
    vscPOLineScroll.Enabled = False
    vscPOLineScroll.Visible = False
    vscPOLineScroll.Min = 1
    vscPOLineScroll.Value = 1
    Exit Sub
  End If

  vscPOLineScroll.Min = 1
  vscPOLineScroll.Max = intPOLines - 9
  vscPOLineScroll.Enabled = True
  vscPOLineScroll.Visible = True
End Sub


Private Sub HighlightPOLineRow(intRowNum As Integer)
  If intHighlightedPOLineRow >= 0 Then
    UnHighlightPOLineRow intHighlightedPOLineRow
  End If
  
  txtItem(intRowNum).BackColor = &HFFFF00
  txtSpace1(intRowNum).BackColor = &HFFFF00
  txtDescription(intRowNum).BackColor = &HFFFF00
  txtSpace2(intRowNum).BackColor = &HFFFF00
  txtQuantity(intRowNum).BackColor = &HFFFF00
  txtSpace3(intRowNum).BackColor = &HFFFF00
  txtRate(intRowNum).BackColor = &HFFFF00
  txtSpace4(intRowNum).BackColor = &HFFFF00
  txtAmount(intRowNum).BackColor = &HFFFF00
  txtReceived(intRowNum).BackColor = &HFFFF00
  txtClosed(intRowNum).BackColor = &HFFFF00

  intHighlightedPOLineRow = intRowNum
  intActualHighlightedPOLine = vscPOLineScroll.Value + intRowNum
End Sub


Private Sub UnHighlightPOLineRow(intRowNum As Integer)
  If intHighlightedPOLineRow >= 0 Then
    txtItem(intRowNum).BackColor = &H80000005
    txtSpace1(intRowNum).BackColor = &H80000005
    txtDescription(intRowNum).BackColor = &H80000005
    txtSpace2(intRowNum).BackColor = &H80000005
    txtQuantity(intRowNum).BackColor = &H80000005
    txtSpace3(intRowNum).BackColor = &H80000005
    txtRate(intRowNum).BackColor = &H80000005
    txtSpace4(intRowNum).BackColor = &H80000005
    txtAmount(intRowNum).BackColor = &H80000005
    txtReceived(intRowNum).BackColor = &H80000005
    txtClosed(intRowNum).BackColor = &H80000005
  End If
  
  intHighlightedPOLineRow = -1
  intActualHighlightedPOLine = -1
End Sub

Private Sub txtAmount_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtClosed_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtDescription_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtItem_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtQuantity_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtRate_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtReceived_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtSpace1_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtSpace2_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtSpace3_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub txtSpace4_Click(Index As Integer)
  If Index < intPOLines Then
    HighlightPOLineRow (Index)
  End If
End Sub

Private Sub vscPOLineScroll_Change()
  UnHighlightPOLineRow intHighlightedPOLineRow
  
  Dim strSplits() As String
  Dim i As Integer
  intTopDisplayedPOLine = vscPOLineScroll.Value
  For i = 0 To 9
    If strPOLines(vscPOLineScroll.Value + i) <> Empty Then
      strSplits = Split(strPOLines(vscPOLineScroll.Value + i), "<spliter>")
      txtItem(i).Text = strSplits(0)
      txtDescription(i).Text = strSplits(1)
      txtQuantity(i).Text = strSplits(2)
      txtRate(i).Text = strSplits(3)
      txtAmount(i).Text = strSplits(5)
      txtReceived(i).Text = strSplits(6)
      txtClosed(i).Text = strSplits(7)
    End If
  Next
  
  If intActualHighlightedPOLine >= vscPOLineScroll.Value And _
     intActualHighlightedPOLine <= vscPOLineScroll.Value + 9 Then
    HighlightPOLineRow intActualHighlightedPOLine - vscPOLineScroll.Value
  End If
  
  If frmPurchaseOrderDetail.Visible Then
    txtItem(0).SetFocus
  End If

End Sub


Private Function POLineClosed(strLine As String) As Boolean

  Dim strSplits() As String
  strSplits = Split(strLine, "<spliter>")
  POLineClosed = (strSplits(7) = "X")
End Function


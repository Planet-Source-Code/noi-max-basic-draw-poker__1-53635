VERSION 5.00
Begin VB.Form frmPoker 
   Caption         =   "Draw Poker"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   7680
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   120
      TabIndex        =   50
      Top             =   5880
      Width           =   3255
      Begin VB.Label Label23 
         Caption         =   "*  Max bet pays an additional 40% on the higher hands."
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Credits:"
      Height          =   3615
      Left            =   3480
      TabIndex        =   31
      Top             =   3360
      Width           =   4095
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   2520
         TabIndex        =   49
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Left            =   2520
         TabIndex        =   48
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         Height          =   255
         Left            =   2520
         TabIndex        =   47
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         Height          =   255
         Left            =   2520
         TabIndex        =   46
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   255
         Left            =   2520
         TabIndex        =   45
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         Height          =   255
         Left            =   2520
         TabIndex        =   43
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "250"
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "One Pair (Jacks or better)"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Two Pair"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Three of a Kind"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Straight "
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Flush "
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Full House*"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Four of a Kind*"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Straight Flush*"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Royal Flush* "
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "How to Play:"
      Height          =   2535
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   3255
      Begin VB.Label Label4 
         Caption         =   "To Stand, double click anywhere on the form to quickly select all the cards and then press ""Draw"""
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Hold cards by clicking the ""Hold"" boxes below the cards you want to keep."
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Choose number of credits to play and press ""Draw"""
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.ListBox List3 
      Height          =   1620
      ItemData        =   "frmCards.frx":0000
      Left            =   5520
      List            =   "frmCards.frx":0002
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "frmCards.frx":0004
      Left            =   4320
      List            =   "frmCards.frx":0006
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMax 
      Caption         =   "Max Bet"
      Height          =   495
      Left            =   6840
      TabIndex        =   24
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "Bet 1"
      Height          =   495
      Left            =   6000
      TabIndex        =   23
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Bet"
      Height          =   855
      Left            =   6120
      TabIndex        =   21
      Top             =   960
      Width           =   1455
      Begin VB.TextBox txtBet 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCredit 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Credits"
      Height          =   855
      Left            =   6120
      TabIndex        =   19
      Top             =   0
      Width           =   1455
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   17
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   16
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "frmCards.frx":0008
      Left            =   3240
      List            =   "frmCards.frx":000A
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   25
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Image imgSpade 
      Height          =   330
      Index           =   4
      Left            =   5280
      Picture         =   "frmCards.frx":000C
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgClub 
      Height          =   330
      Index           =   4
      Left            =   5280
      Picture         =   "frmCards.frx":0196
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgHeart 
      Height          =   330
      Index           =   4
      Left            =   5280
      Picture         =   "frmCards.frx":0320
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDiamond 
      Height          =   330
      Index           =   4
      Left            =   5280
      Picture         =   "frmCards.frx":04AA
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5160
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgSpade 
      Height          =   330
      Index           =   3
      Left            =   4080
      Picture         =   "frmCards.frx":0634
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgClub 
      Height          =   330
      Index           =   3
      Left            =   4080
      Picture         =   "frmCards.frx":07BE
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgHeart 
      Height          =   330
      Index           =   3
      Left            =   4080
      Picture         =   "frmCards.frx":0948
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDiamond 
      Height          =   330
      Index           =   3
      Left            =   4080
      Picture         =   "frmCards.frx":0AD2
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgSpade 
      Height          =   330
      Index           =   1
      Left            =   1680
      Picture         =   "frmCards.frx":0C5C
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgClub 
      Height          =   330
      Index           =   1
      Left            =   1680
      Picture         =   "frmCards.frx":0DE6
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgHeart 
      Height          =   330
      Index           =   1
      Left            =   1680
      Picture         =   "frmCards.frx":0F70
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDiamond 
      Height          =   330
      Index           =   1
      Left            =   1680
      Picture         =   "frmCards.frx":10FA
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDiamond 
      Height          =   330
      Index           =   0
      Left            =   480
      Picture         =   "frmCards.frx":1284
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgHeart 
      Height          =   330
      Index           =   0
      Left            =   480
      Picture         =   "frmCards.frx":140E
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgClub 
      Height          =   330
      Index           =   0
      Left            =   480
      Picture         =   "frmCards.frx":1598
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgSpade 
      Height          =   330
      Index           =   0
      Left            =   480
      Picture         =   "frmCards.frx":1722
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgDiamond 
      Height          =   330
      Index           =   2
      Left            =   2880
      Picture         =   "frmCards.frx":18AC
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgHeart 
      Height          =   330
      Index           =   2
      Left            =   2880
      Picture         =   "frmCards.frx":1A36
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgClub 
      Height          =   330
      Index           =   2
      Left            =   2880
      Picture         =   "frmCards.frx":1BC0
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSpade 
      Height          =   330
      Index           =   2
      Left            =   2880
      Picture         =   "frmCards.frx":1D4A
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   4
      Left            =   4920
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnucredits 
         Caption         =   "Credits"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmPoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ':)

' ********************************************************************************
'*                                                                                *
'*                                                                                *
'*  DRAW POKER! By noi_max                                                        *
'*  03/15/2004  Version 1.0                                                       *
'*                                                                                *
'*  This is my first game and is heavily commented.                               *
'*  I wish I could remember where the shuffle algorithm                           *
'*  came from :)                                                                  *
'*  Credit for the Bubble Sort goes to Squirm from                                *
'*  his tutorial.                                                                 *
'*                                                                                *
'*  DEPENDENCY INFORMATION-                                                       *
'*  This card game uses the suit images from the default                          *
'*  bitmap folder installed with VB6 using the path                               *
'*  C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\Assorted\    *
'*                                                                                *
'*  If you used the default folders when installing VB6 it should find the        *
'*  images with no problems.                                                      *
'*                                                                                *
'*                                                                                *
'*  www.visualbasicforum.com                                                      *
'*                                                                                *
' ********************************************************************************

Private Sub Form_Load()

Me.Top = 0
Me.Height = 3735

Shuffle1 'go ahead and shuffle on load
txtCredit.Text = 50
txtBet.Text = 1


End Sub

Private Sub Form_DblClick()
'double click the form to select all the hold boxes

Dim ctrl As Control

For Each ctrl In Me.Controls        'loop through the controls and change the
   If TypeOf ctrl Is CheckBox Then  'values of the check boxes
      ctrl.Value = 1
   End If
Next

End Sub

Private Sub cmdOne_Click()

Static i As Integer

'This will add one to the current bet but not allow for
'a bet more than the max (10)

i = 1
i = CInt(txtBet.Text) + 1
If i > 10 Then
   i = 1
End If

If Bet > Credit Then
   Bet = 1
End If

txtBet.Text = CStr(i)

End Sub

Private Sub cmdMax_Click()

txtBet.Text = 10 'max bet

End Sub

Public Sub Shuffle1()
'I don't remember whom to credit for the shuffle algorithm :(
List1.Clear

Dim Deck(1 To 52) As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim temp As Long

For i = 1 To 52
Deck(i) = i
Next i

Randomize Timer
For i = 1 To 10
For j = 1 To 52
k = (Int(Rnd * 52) + 1)
temp = Deck(j)
Deck(j) = Deck(k)
Deck(k) = temp
Next j
Next i

For i = 1 To 52
List1.AddItem Deck(i) 'add our shuffled cards to a listbox.
Next i
End Sub

Private Sub cmdDraw_Click()
Static i As Integer
Dim j As Integer

'This block sets up which turn it is.
i = i + 1

If i > 1 Then                    'if it's the second turn
   For j = 0 To 4
      chkHold(j).Enabled = False 'turn off the hold checkboxes
   Next j
   i = 0
End If

'clear list2 at beginning of first turn only
If i = 1 Then
   List2.Clear
End If

lblScore.Caption = ""
lblScore.BackColor = &H8000000F
DrawCards (i)  'turn number

End Sub

Public Sub DrawCards(Index As Integer) 'index indicates what turn number.
Dim i As Integer
Dim j As Integer
Dim Suit As String
Dim Val As String
Dim intResponse As Integer

      cmdOne.Enabled = False  'after we make the first draw disable the bets.
      cmdMax.Enabled = False

Credit = CLng(txtCredit.Text) 'set module level variables
Bet = CLng(txtBet.Text)

If Credit < 10 And Credit > 1 Then 'we can't still bet 10 if we have
   Bet = Credit                    'less than 10 credits.
   txtBet.Text = CStr(Bet)
End If

If Index = 1 Then                 'if it's the first turn
   Shuffle1                       'shuffle the cards
   Credit = Credit - Bet          'take away bet amount from credits
   txtCredit.Text = CStr(Credit)  're-display remaining credits
   
   For j = 0 To 4
      chkHold(j).Value = 0        'turn the hold checkboxes off
   Next j
End If

For j = 0 To 4
  If chkHold(j).Value = 0 Then 'if we're not holding any cards
    If Index = 1 Then          'if it's our first turn
    chkHold(j).Enabled = True  'turn on the check boxes
    End If
    
    'make the cards disappear so they can later reappear
    lblVal(j).Caption = ""
    lblVal(j).Visible = False
    Label1(j).Visible = False
    imgSpade(j).Visible = False
    imgClub(j).Visible = False
    imgHeart(j).Visible = False
    imgDiamond(j).Visible = False
   End If
Next j

If List1.ListCount < 5 Then Exit Sub 'if there's less than 5 cards left in the deck

 For i = 0 To 4  '<===BEGINNING OF MAIN LOOP=====>
    
    If chkHold(i).Value = 0 Then  'if we're not holding any cards
       cmdDraw.Enabled = False
    
       Wait 10 '(Module function) allows time for the cards to appear one by one.
     
       Text1.Text = CStr(List1.List(0))
     
       If Index = 1 Then
          List2.AddItem Text1.Text 'add five drawn cards to List2.
       End If
     
     If Index = 0 And chkHold(i).Value = 0 Then 'second turn
        List2.RemoveItem (i)          'take the cards we didn't hold off the list
        List2.AddItem Text1.Text, (i) 'put the new cards in their place
        
     End If
     
     List1.RemoveItem 0              'take the card off the first list
                                     'so it cannot be drawn again from the pack
                                     
     Suit = Card(CInt(Text1.Text))   'module function to determine suit.
     Val = Value(CInt(Text1.Text))   'module function to determine card value
     
     Label1(i).Visible = True        'show the card and the value
     lblVal(i).Visible = True
     lblVal(i).Caption = Val
     
     'Conditional block to handle which images appear for suit.
     If Suit = "Spades" Then
        imgSpade(i).Visible = True
     Else
        imgSpade(i).Visible = False
     End If
     
     If Suit = "Clubs" Then
        imgClub(i).Visible = True
     Else
        imgClub(i).Visible = False
     End If
     
     If Suit = "Hearts" Then
        imgHeart(i).Visible = True
     Else
        imgHeart(i).Visible = False
     End If
     
     If Suit = "Diamonds" Then
        imgDiamond(i).Visible = True
     Else
        imgDiamond(i).Visible = False
     End If
     
     cmdDraw.Enabled = True
   End If

Next i  '<====END OF MAIN LOOP====>

If Index = 0 Then
   lblScore.BackColor = &H8080FF
   lblScore.Caption = "GAME OVER"

   Score 'Sub procedure below this one to see if we won anything.
     
End If
 

'If we run out of credits it's game over man!!!!!
If Credit = 0 And Index = 0 And blnWin = False Then
   intResponse = MsgBox("You have run out of Credits!" & vbCrLf & _
   "Would you like to start a new game?", vbOKCancel, "Draw Poker")
   If intResponse = vbOK Then
      Unload Me
      Load frmPoker
      frmPoker.Show
   Else
      Unload Me
      End
   End If
   
End If

'If the second turn is finished, put the bets back on for the next draw.
If Index = 0 Then
   cmdOne.Enabled = True
   cmdMax.Enabled = True
End If

End Sub

Public Sub Score()
'This procedure is waaaay too long. Maybe someone could optimize it for me?? :)
'"Win" is a module procedure to beep the computer and display what we've got. It's
'used a lot below.

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim int2(0 To 12, 0 To 3) As Integer
Dim arCards() As Integer
Dim arStraight() As String
Dim arFlush() As String
Dim Multi(12) As Integer
Dim Pair As Integer
Dim Sequence As Integer
Dim Flush As Integer
Dim blnTriple As Boolean
Dim blnPair As Boolean
Dim blnFour As Boolean
Dim blnRoyal As Boolean
Dim blnFlush As Boolean
Dim blnStraight As Boolean

blnWin = False 'module variable: start this off as false until the Win sub
               'tells us otherwise.
               
'set up a 2 dimensional array to represent all 52 cards and later sort them
'by value

int2(0, 0) = 1
int2(0, 1) = 14
int2(0, 2) = 27
int2(0, 3) = 40

int2(1, 0) = 2
int2(1, 1) = 15
int2(1, 2) = 28
int2(1, 3) = 41

int2(2, 0) = 3
int2(2, 1) = 16
int2(2, 2) = 29
int2(2, 3) = 42

int2(3, 0) = 4
int2(3, 1) = 17
int2(3, 2) = 30
int2(3, 3) = 43

int2(4, 0) = 5
int2(4, 1) = 18
int2(4, 2) = 31
int2(4, 3) = 44

int2(5, 0) = 6
int2(5, 1) = 19
int2(5, 2) = 32
int2(5, 3) = 45

int2(6, 0) = 7
int2(6, 1) = 20
int2(6, 2) = 33
int2(6, 3) = 46

int2(7, 0) = 8
int2(7, 1) = 21
int2(7, 2) = 34
int2(7, 3) = 47

int2(8, 0) = 9
int2(8, 1) = 22
int2(8, 2) = 35
int2(8, 3) = 48

int2(9, 0) = 10
int2(9, 1) = 23
int2(9, 2) = 36
int2(9, 3) = 49

int2(10, 0) = 11
int2(10, 1) = 24
int2(10, 2) = 37
int2(10, 3) = 50

int2(11, 0) = 12
int2(11, 1) = 25
int2(11, 2) = 38
int2(11, 3) = 51

int2(12, 0) = 13
int2(12, 1) = 26
int2(12, 2) = 39
int2(12, 3) = 52

ReDim arCards(4) 'array for 5 cards
ReDim arStraight(4)
ReDim arFlush(4)

'Get the cards from our list
For i = 0 To 4
   arCards(i) = List2.List(i)
Next i

'Call the Card Function to determine suit (it's in the module!)
For i = 0 To 4
   arFlush(i) = Card(arCards(i))
Next i

'Compare the cards and see how many are the same suit
For i = 1 To 4
   If arFlush(0) = arFlush(i) Then
      Flush = Flush + 1
   End If
Next i

'If they're all the same suit set the boolean to true
If Flush = 4 Then
   blnFlush = True
End If

'Get the alpha values and convert them to numbers for sorting
For i = 0 To 4
   arStraight(i) = Value(arCards(i))
   If arStraight(i) = "J" Then
      arStraight(i) = "11"
   ElseIf arStraight(i) = "Q" Then
      arStraight(i) = "12"
   ElseIf arStraight(i) = "K" Then
      arStraight(i) = "13"
   ElseIf arStraight(i) = "A" Then
      arStraight(i) = "1"
   End If
Next i


'Sort the numeric values (Bubblesort courtesy of Squirm's tutorial)
BubbleSort arStraight '(that's the Public Sub below this one)

'Determine if we have a straight of Royal value; if so, set a boolean to true
If arStraight(0) = "1" And _
   arStraight(1) = "10" And _
   arStraight(2) = "11" And _
   arStraight(3) = "12" And _
   arStraight(4) = "13" Then
   blnRoyal = True
End If
   
'This will add 1 to Sequence for the number of cards that are, well, in sequence. :)
For j = 1 To 4
   If CInt(arStraight(0)) = CInt(arStraight(j)) - j Then
      Sequence = Sequence + 1
      End If
Next j

'This loops through the array to find cards of matching values. When a match is
'found it will add one to the integer Multi to determine if there's 2 or more
'of the same card present. This is used for 2, 3 and 4 of a kind and sets
'up the other conditional loops.

For i = 0 To 4
   For j = 0 To 12
      For k = 0 To 3
         If CInt(arCards(i)) = int2(j, k) Then
            Multi(j) = Multi(j) + 1
         End If
      Next k
   Next j
Next i


'This loop sets up a variable called Pair to determine if there's more
'than 1 pair of like values (two pair)

For i = 0 To 12
   If Multi(i) > 1 And Multi(i) < 3 Then
      Pair = Pair + 1
   End If
Next i

'Declare a win on 2 pair and set a boolean used for a condition to
'avoid also hitting Jacks or Better.

If Pair >= 2 Then
   Win ("Two Pair!"), (2) 'module procedure with 2 arguments
   blnPair = True
End If

'This loop structure will check for 3 or 4 of a kind and set up
'boolean values for later comparison.

For i = 0 To 12
   If Multi(i) > 2 And Multi(i) < 4 Then 'three of a kind
      blnTriple = True
   End If
   
   If Multi(i) > 3 Then 'four of a kind
      blnFour = True
   End If
Next i


'If we have 4 in the variable "Sequence" then we have a straight!
If Sequence = 4 Then
   blnStraight = True
End If

'If our Royal and Flush booleans are true well....
If blnRoyal = True And blnFlush = True Then
   If Bet = 10 Then
      Win ("Royal Flush!!"), (350) '40% more for max bet
   Else
      Win ("Royal Flush!!"), (250)
   End If
   Exit Sub
End If

'Here we check for a straight flush
If blnStraight = True And blnFlush = True Then
   If Bet = 10 Then
      Win ("Straight Flush!!"), (70) '40% more for max bet
   Else
      Win ("Straight Flush!!"), (50)
   End If
   Exit Sub
End If

'Just a plain ol' flush
If blnFlush = True Then
   Win ("Flush!"), (6)
   Exit Sub
End If

'Just a plain ol' straight
If blnStraight = True Or blnRoyal = True And blnFlush = False Then
   Win ("Straight!"), (4)
   Exit Sub
End If


'If theres a 3 of a kind and a pair then...
If blnTriple = True And Pair = 1 Then
   If Bet = 10 Then
      Win ("Full House!"), (12) 'almost 40% more for max bet
   Else
      Win ("Full House!"), (9)
   End If
   Exit Sub
End If

'If there's only a three of a kind...
If blnTriple = True And Pair <> 1 And blnFour = False Then
   Win ("Three of a Kind!"), (3)
   Exit Sub
End If

'If there's a four of a kind...
If blnFour = True Then
   If Bet = 10 Then
      Win ("Four of a Kind!"), (35) '40% more for max bet
   Else
      Win ("Four of a Kind!"), (25)
   End If
   Exit Sub
End If

'If it's only Jacks or Better...
If blnTriple = False And blnPair = False And blnFour = False Then
    For i = 9 To 12
       If Multi(i) > 1 Then
          Win ("Jacks or Better"), (1)
       End If
    Next i
End If

End Sub

Public Sub BubbleSort(intArray() As String)
Dim iOuter As Integer
Dim iInner As Integer
Dim iLBound As Integer
Dim iUBound As Integer
Dim iTemp As Integer
Dim i As Integer

iLBound = LBound(intArray)
iUBound = UBound(intArray)

'Which bubbling pass
For iOuter = iLBound To iUBound - 1
   'Which comparison
   For iInner = iLBound To iUBound - iOuter - 1

      'Compare this item to the next item
      If CInt(intArray(iInner)) > CInt(intArray(iInner + 1)) Then
         'Swap
         iTemp = CInt(intArray(iInner))
         intArray(iInner) = intArray(iInner + 1)
         intArray(iInner + 1) = iTemp
      End If

   Next iInner
Next iOuter

End Sub

Private Sub mnuabout_Click()

MsgBox "Draw Poker v 1.0" & vbCrLf & "by noi_max", vbOKOnly, "Draw Poker"
'my shameless self promotion

End Sub

Private Sub mnucredits_Click()

Static i As Integer 'use i as my toggle switch
i = i + 1
If i > 1 Then
   i = 0
   Me.Height = 3735 'show only the main game
   mnucredits.Checked = False
End If

If i = 1 Then
   Me.Height = 7785 'show the credits and basic help
   mnucredits.Checked = True
End If

End Sub

Private Sub mnuExit_Click()

Unload Me  'bye!

End Sub

Private Sub mnuNew_Click()

Unload Me
Load frmPoker
frmPoker.Show 'reload the form for a new game.

End Sub

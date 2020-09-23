Attribute VB_Name = "modCards"
Option Explicit

Public Bet As Long
Public Credit As Long
Public blnWin As Boolean


Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'WAIT Function. tell the sub to wait by entering: Wait 1 (# of seconds)
Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait * 50 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime       '^modified
    DoEvents
Loop
End Function

Public Function Card(Index As Integer) As String
'select case to determine suit

Dim Suit As String

Select Case Index
   Case Is >= 40
      Suit = "Clubs"
   Case Is >= 27
      Suit = "Spades"
   Case Is >= 14
      Suit = "Diamonds"
   Case Is >= 1
      Suit = "Hearts"
End Select
   
Card = Suit
   
End Function

Public Function Value(Index As Integer) As String
'If Then ElseIf to determine face cards

Dim Val As String

If Index = 10 Or Index = 23 Or Index = 36 Or Index = 49 Then
      Val = "J"
ElseIf Index = 11 Or Index = 24 Or Index = 37 Or Index = 50 Then
      Val = "Q"
ElseIf Index = 12 Or Index = 25 Or Index = 38 Or Index = 51 Then
      Val = "K"
ElseIf Index = 13 Or Index = 26 Or Index = 39 Or Index = 52 Then
      Val = "A"

ElseIf Index = 1 Or Index = 14 Or Index = 27 Or Index = 40 Then
   Val = "2"
ElseIf Index = 2 Or Index = 15 Or Index = 28 Or Index = 41 Then
   Val = "3"
ElseIf Index = 3 Or Index = 16 Or Index = 29 Or Index = 42 Then
   Val = "4"
ElseIf Index = 4 Or Index = 17 Or Index = 30 Or Index = 43 Then
   Val = "5"
ElseIf Index = 5 Or Index = 18 Or Index = 31 Or Index = 44 Then
   Val = "6"
ElseIf Index = 6 Or Index = 19 Or Index = 32 Or Index = 45 Then
   Val = "7"
ElseIf Index = 7 Or Index = 20 Or Index = 33 Or Index = 46 Then
   Val = "8"
ElseIf Index = 8 Or Index = 21 Or Index = 34 Or Index = 47 Then
   Val = "9"
ElseIf Index = 9 Or Index = 22 Or Index = 35 Or Index = 48 Then
   Val = "10"
End If

Value = Val
   
End Function


Public Sub Win(Argument As String, Figure As Long)

blnWin = True 'this is put here in case you win but have no credits left,
              'otherwise the game would end prematurely.
              
frmPoker.cmdDraw.Enabled = False 'Don't allow the user to hit "Draw" while the
                                 'credits are ringing up.
Dim i As Long

For i = 1 To (Bet * Figure)
      Beep
      Credit = Credit + 1
      frmPoker.txtCredit.Text = CStr(Credit)
      frmPoker.lblScore.BackColor = &HFFFFC0
      frmPoker.lblScore.Caption = "WIN  " & Argument
      Wait 1 'this is from the GetTickCount above..
Next i

frmPoker.cmdDraw.Enabled = True

End Sub



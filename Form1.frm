VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SysTrayText"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4680
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Remove"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter your text here and press enter"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SysTrayText by BIG BAD BOOTZ
' © BIG BAD BOOTZ programming 2001
' http://surf.to/bootz
'
' not much comments, too lame to write them :P
' you can change how many characters are displayed at once
' by adjusting the size variable at the end of the form_load sub
'
' i'll add more functions later (like sending eachother systray
' messages over a network or over the internet), but you can do
' that yourself too :P

Dim char$(68)
Dim char2$(68)
Dim frm(25) As New frmIcon
Dim scrolltext$
Dim scrollpos
Dim size

Private Sub Command1_Click()
scrolltext$ = ""
scrollpos = 0
Timer1.Interval = 0
Timer1.Enabled = False
If Len(Text1.Text) > size Then
scrolltext$ = Text1.Text
scrollpos = 1
Text1.Text = Left$(scrolltext$, size)
End If
For a = 0 To 25
Unload frm(a)
Next a
For a = 1 To Len(Text1.Text)
found = 0
For b = 1 To 68
If LCase$(Mid$(Text1.Text, a, 1)) = LCase$(char2$(b)) Then found = b: Exit For
Next b
If found = 0 Then found = 68
f$ = App.Path & "\" & char$(found) & ".ico"
frm(a).Icon = LoadPicture(f$)
frm(a).Timer1.Enabled = True
frm(a).Timer1.Interval = 100 + (a * 50)
Next a
If scrolltext <> "" Then
Text1.Text = scrolltext$
scrolltext$ = scrolltext$ & " --- "
Timer1.Enabled = True
Timer1.Interval = 3000
End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer1.Interval = 0
For a = 0 To 25
Unload frm(a)
Next a
End Sub

Private Sub Form_Load()
char$(1) = "a"
char$(2) = "b"
char$(3) = "c"
char$(4) = "d"
char$(5) = "e"
char$(6) = "f"
char$(7) = "g"
char$(8) = "h"
char$(9) = "i"
char$(10) = "j"
char$(11) = "k"
char$(12) = "l"
char$(13) = "m"
char$(14) = "n"
char$(15) = "o"
char$(16) = "p"
char$(17) = "q"
char$(18) = "r"
char$(19) = "s"
char$(20) = "t"
char$(21) = "u"
char$(22) = "v"
char$(23) = "w"
char$(24) = "x"
char$(25) = "y"
char$(26) = "z"
char$(27) = "0"
char$(28) = "1"
char$(29) = "2"
char$(30) = "3"
char$(31) = "4"
char$(32) = "5"
char$(33) = "6"
char$(34) = "7"
char$(35) = "8"
char$(36) = "9"
char$(37) = "("
char$(38) = ")"
char$(39) = "["
char$(40) = "]"
char$(41) = "accolade-open"
char$(42) = "accolade-close"
char$(43) = "and"
char$(44) = "asterisk"
char$(45) = "at"
char$(46) = "backslash"
char$(47) = "caret"
char$(48) = "colon"
char$(49) = "dollar"
char$(50) = "dot"
char$(51) = "double-quotes"
char$(52) = "equals"
char$(53) = "exclamation"
char$(54) = "greater-than"
char$(55) = "komma"
char$(56) = "less-than"
char$(57) = "line"
char$(58) = "minus"
char$(59) = "percent"
char$(60) = "plus"
char$(61) = "pound"
char$(62) = "questionmark"
char$(63) = "semicolon"
char$(64) = "single-quote"
char$(65) = "slash"
char$(66) = "tilde"
char$(67) = "underscore"
char$(68) = "space"
For a = 1 To 40
char2$(a) = char$(a)
Next a
char2$(41) = "{"
char2$(42) = "}"
char2$(43) = "&"
char2$(44) = "*"
char2$(45) = "@"
char2$(46) = "\"
char2$(47) = "^"
char2$(48) = ":"
char2$(49) = "$"
char2$(50) = "."
char2$(51) = Chr$(34)
char2$(52) = "="
char2$(53) = "!"
char2$(54) = ">"
char2$(55) = ","
char2$(56) = "<"
char2$(57) = "|"
char2$(58) = "-"
char2$(59) = "%"
char2$(60) = "+"
char2$(61) = "#"
char2$(62) = "?"
char2$(63) = ";"
char2$(64) = "'"
char2$(65) = "/"
char2$(66) = "~"
char2$(67) = "_"
char2$(68) = " "
size = 12 'do NOT set over 25!!!!
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Command2_Click
End
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval <> 200 Then Timer1.Interval = 200
scrollpos = scrollpos + 1
If scrollpos > Len(scrolltext$) Then scrollpos = 1
If scrollpos + size > Len(scrolltext$) Then
s$ = Right$(scrolltext$, size - (scrollpos + size - Len(scrolltext$)))
s$ = s$ & Left$(scrolltext$, scrollpos + size - Len(scrolltext$))
Else
s$ = Mid$(scrolltext$, scrollpos, size)
End If
For a = 1 To size
found = 0
For b = 1 To 68
If LCase$(Mid$(s$, a, 1)) = LCase$(char2$(b)) Then found = b: Exit For
Next b
If found = 0 Then found = 68
f$ = App.Path & "\" & char$(found) & ".ico"
frm(a).Icon = LoadPicture(f$)
frm(a).Timer2.Enabled = True
frm(a).Timer2.Interval = 1
Next a
End Sub

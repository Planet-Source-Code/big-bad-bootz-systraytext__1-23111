VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SysTrayText by BIG BAD BOOTZ
'
'most of this code was taken from microsoft:
'http://support.microsoft.com/support/kb/articles/q162/6/13.asp?LN=EN-US&SD=gn&FR=0&qry=tray icon&rnk=2&src=DHCS_MSPSS_gn_SRCH&SPR=VBB
'there are more examples at their site
'use this url for a search:
'http://search.support.microsoft.com/kb/psssearch.asp?SPR=vbb&T=B&KT=ALL&T1=7d&LQ=tray+icon&PQ=0&S=F&A=T&DU=C&FR=0&D=vbwin&LPR=&LNG=ENG&VR=http://support.microsoft.com/support;http://support.microsoft.com/servicedesks/webcasts;http://support.microsoft.com/highlights&CAT=Support&VRL=ENG&SA=GN&Go.x=10&Go.y=13
'
'!!!! i removed some things from the original code so go visit the site
'     mentioned above if you want the original code with more functions
'     (mouse functions for popup-menu and that kind of things)


      'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      Private Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
      Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Private Const WM_LBUTTONDOWN = &H201     'Button down
      Private Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Private Const WM_RBUTTONDOWN = &H204     'Button down
      Private Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Dim nid As NOTIFYICONDATA

      Private Sub add()
         'Click this button to add an icon to the taskbar status area.

         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = Me.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = vbNull
         nid.hIcon = Me.Icon
         nid.szTip = "SysTrayText" & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid
      End Sub
      
      Private Sub change()
         'Click this button to add an icon to the taskbar status area.

         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = Me.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = vbNull
         nid.hIcon = Me.Icon
         nid.szTip = "SysTrayText" & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_MODIFY, nid
      End Sub

      Private Sub remove()
         'Click this button to delete the added icon from the taskbar
         'status area by calling the Shell_NotifyIcon function.
         Shell_NotifyIcon NIM_DELETE, nid
      End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
         'Delete the added icon from the taskbar status area when the
         'program ends.
         Shell_NotifyIcon NIM_DELETE, nid
End Sub

      Private Sub Form_Terminate()
         'Delete the added icon from the taskbar status area when the
         'program ends.
         Shell_NotifyIcon NIM_DELETE, nid
      End Sub

'we need these two timers to call the right subroutines from an other form
Private Sub Timer1_Timer()
Timer1.Enabled = False
Timer1.Interval = 0
Call add
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer2.Interval = 0
Call change
End Sub

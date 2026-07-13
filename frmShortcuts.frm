VERSION 5.00
Begin VB.Form frmShortcuts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modernizr - Keyboard Shortcuts"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstKeys 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6440
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   360
      Left            =   5500
      TabIndex        =   1
      Top             =   5070
      Width           =   1120
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Cheat sheet for everything the add-in binds. Ctrl+Shift+/ in a code
' window, or Modernizr -> Keyboard Shortcuts.

Public Sub ShowSheet()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Me.Show vbModeless
End Sub

Private Sub Form_Load()
    Populate
End Sub

Private Sub Populate()
    lstKeys.Clear
    Hdr "FIND"
    row "Ctrl+F", "Show / hide find / replace bar"
    row "Ctrl+Shit+F", "Find in files"
    row "Enter / Shift+Enter", "Find next / previous (in the bar)"
    row "F3 / Shift+F3", "Find next / previous (anywhere)"
    row "Ctrl+F3", "Highlight word at cursor everywhere"
    row "Esc", "Close find bar"
    Hdr ""
    Hdr "NAVIGATE  (focus in code window)"
    row "F12", "Go to definition of word at cursor"
    row "Shift+F12", "Find all references to word at cursor"
    row "Ctrl+F2", "Toggle bookmark on current line"
    row "F2", "Next bookmark"
    row "Ctrl+Shift+O", "Code browser (procedures / TODOs)"
    row "Ctrl+P", "Quick Open: fuzzy find file / module"
    row "", "  (Enter = code, Shift+Enter = designer)"
    row "Alt+Left / Alt+Right", "Back / forward (navigation history)"
    Hdr ""
    Hdr "EDIT  (focus in code window)"
    row "Ctrl+D", "Duplicate line / selection"
    row "Alt+Up / Alt+Down", "Move lines up / down"
    row "Ctrl+Shift+K", "Delete lines"
    row "Ctrl+/", "Comment / uncomment"
    Hdr ""
    Hdr "WINDOWS"
    row "Ctrl+Tab", "Switch window (hold Ctrl, Tab cycles,"
    row "", "  release Ctrl to switch, Esc cancels)"
    row "Ctrl+Shift+Tab", "Switch window, backwards"
    row "Middle-click tab", "Close window"
    row "Drag tab", "Reorder tabs"
    row "Right-click tab", "Close / Copy Path / Open Folder..."
    row "Esc", "Close the active Modernizr window"
    Hdr ""
    Hdr "GIT  (repo auto-detected from the project folder)"
    row "Ctrl+Shift+G", "Changes: stage / unstage / commit"
    row "Ctrl+Shift+L", "Log with branch graph"
    row "Ctrl+Shift+B", "Blame current line"
    row "", "Tab bar shows branch; margin bars show"
    row "", "  changed lines (green=add, blue=edit)"
    Hdr ""
    Hdr "MOUSE"
    row "Wheel", "Scroll code window"
    row "Shift+Wheel", "Scroll horizontally"
    Hdr ""
    Hdr "HELP"
    row "Ctrl+Shift+/", "This window"
End Sub

Private Sub Hdr(ByVal s As String)
    lstKeys.AddItem s
End Sub

Private Sub row(ByVal keys As String, ByVal desc As String)
    lstKeys.AddItem "  " & Left$(keys & Space$(22), 22) & desc
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

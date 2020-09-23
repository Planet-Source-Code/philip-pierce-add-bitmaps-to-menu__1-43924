VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   360
      Picture         =   "Form1.frx":0372
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo err
    Dim hMenu As Long, hSubmenu As Long
    Dim hID As Long
    
    Me.Hide
    
    'Get the menuhandle of your app
    hMenu = GetMenu(Me.hwnd)
    
    'Get the handle of the first submenu (Hello)
    hSubmenu = GetSubMenu(hMenu, 0)
    
    'Get the menuId of the first entry (Bitmap)
    hID = GetMenuItemID(hSubmenu, 0)
    
    'Add the bitmap
    'You can add two bitmaps to a menuentry
    'One for the checked and one for the unchecked
    'state.
    SetMenuItemBitmaps hMenu, hID, MF_BITMAP, Me.Picture1.Picture, Me.Picture2.Picture
    
    ' repop the picture box
    Set Me.Picture3.Picture = LoadPicture(App.Path & "\lake.bmp")
    
    ' do the next submenu
    hID = GetMenuItemID(hSubmenu, 1)
    
    ' add the bitmap
    SetMenuItemBitmaps hMenu, hID, MF_BITMAP, Me.Picture3.Picture, Me.Picture3.Picture
    
    Me.Show
    Exit Sub
err:
    err.Clear
    Me.Show
End Sub

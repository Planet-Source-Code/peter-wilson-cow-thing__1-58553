VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmFarm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   Icon            =   "frmFarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1620
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox imgMooing 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   3060
      ScaleHeight     =   96
      ScaleMode       =   0  'User
      ScaleWidth      =   329.143
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox imgWalking 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   1560
      ScaleHeight     =   96
      ScaleMode       =   0  'User
      ScaleWidth      =   329.143
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox imgEating 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   60
      ScaleHeight     =   96
      ScaleMode       =   0  'User
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Timer CowAnimation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3660
      Top             =   1680
   End
End
Attribute VB_Name = "frmFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ====================================================
' Cow Thing v1.0
' Copyright © 2005 Peter Wilson - All rights reserved.
'
' Also see the included readme.txt file.
' ====================================================

' Cow Behaviour
Enum CowState
    CowWalking
    CowMooing
    CowEatingStart
    CowEating
    CowEatingFinish
End Enum


' Simple Cow data type.
Private Type Cow
    Enabled             As Boolean
    XPosition           As Single
    YPosition           As Single
    CowSpeed            As Single   ' We can't really change the speed too much since the animation is fixed.
    State               As CowState
    StateIndex          As Integer
    Size                As Single
End Type


' Create an array of Cows
Private m_Cows(1000) As Cow



Private Sub DoLoadCowEating()

    Dim intN            As Integer
    Dim strFileName     As String
    
    
    ' =====================================================================
    ' First control (imgEating) already exists, so don't try to load it.
    ' It has already been loaded by VB.
    ' =====================================================================
    imgEating(0).Picture = LoadPicture(App.Path & "\eating e" & Format(0, "0000") & ".gif")
    
    ' =========================================================
    ' Loop through files, and load new 'imgEating' controls.
    ' =========================================================
    For intN = 1 To 8
        
        ' ===========================================================
        ' Load a new control into the control-array 'imgEating(n)'
        ' ===========================================================
        Load imgEating(intN)
        
        
        ' ================================
        ' Calculate the correct file name.
        ' ================================
        strFileName = App.Path & "\eating e" & Format(intN, "0000") & ".gif"
        
        
        ' ============================================
        ' Load the picture from the applications path.
        ' ============================================
        imgEating(intN).Picture = LoadPicture(strFileName)
        
    Next intN
    
End Sub

Private Sub DoLoadCowMooing()

    ' For comments, see the similar 'DoLoadCowEating' routine.
    Dim intN            As Integer
    Dim strFileName     As String
    imgMooing(0).Picture = LoadPicture(App.Path & "\muuuh e" & Format(0, "0000") & ".gif")
    For intN = 1 To 10
        Load imgMooing(intN)
        strFileName = App.Path & "\muuuh e" & Format(intN, "0000") & ".gif"
        imgMooing(intN).Picture = LoadPicture(strFileName)
        imgMooing(intN).AutoRedraw = True
    Next intN
    
End Sub

Private Sub DoLoadCowWalking()

    ' For comments, see the similar 'DoLoadCowEating' routine.
    Dim intN            As Integer
    Dim strFileName     As String
    imgWalking(0).Picture = LoadPicture(App.Path & "\walking e" & Format(0, "0000") & ".gif")
    For intN = 1 To 7
        Load imgWalking(intN)
        strFileName = App.Path & "\walking e" & Format(intN, "0000") & ".gif"
        imgWalking(intN).Picture = LoadPicture(strFileName)
        imgWalking(intN).AutoRedraw = True
        
    Next intN
    
End Sub


Private Sub CowAnimation_Timer()

    Call DoAnimateCows

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    
End Sub

Private Sub DoModifyFormProperties()

    Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)

    ' Make this form transparent.
    bTrans = 255
    lOldStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        
    ' Transparent Form
    Me.BackColor = RGB(255, 0, 0)

    ' Ghost Cows!
    SetLayeredWindowAttributes Me.hwnd, RGB(255, 0, 0), bTrans, LWA_ALPHA Or LWA_COLORKEY
    
End Sub

Private Sub Form_Load()

    Call DoModifyFormProperties

    
  ' Set properties needed by MCI to open.
   MMControl1.Notify = False
   MMControl1.Wait = False
   MMControl1.Shareable = False
   MMControl1.DeviceType = "WaveAudio"
   MMControl1.FileName = App.Path & "\cow1.wav"
   MMControl1.Command = "Open"
   MMControl1.Command = "Play"
   
    
    ' ========================
    ' Load the Cow Animations.
    ' ========================
    Call DoLoadCowEating
    Call DoLoadCowMooing
    Call DoLoadCowWalking
    
    Me.Show
    DoEvents
    
    Dim cowIndex As Integer
    For cowIndex = LBound(m_Cows) To UBound(m_Cows)
        m_Cows(cowIndex).Enabled = False                    ' Make all cows (except one enabled - ha ha)
        m_Cows(cowIndex).XPosition = Rnd * Me.ScaleWidth
        m_Cows(cowIndex).YPosition = Rnd * Me.ScaleHeight
        m_Cows(cowIndex).State = Int(Rnd * 3)
        m_Cows(cowIndex).StateIndex = Int(Rnd * 7)
        m_Cows(cowIndex).CowSpeed = 3 + Rnd * 4
        m_Cows(cowIndex).Size = Abs(Rnd + 1)
    Next cowIndex
    m_Cows(0).Enabled = True
    
    
    Me.CowAnimation.Enabled = True
    
End Sub

Public Sub DoAnimateCows()

    Dim cowIndex As Integer
    Dim intRND As Integer
    Dim varArray() As Variant

    Dim hdcWidth As Long
    Dim hdcHeight As Long
    Dim hdcWidth2 As Long
    Dim hdcHeight2 As Long
        
    hdcWidth = Me.imgEating(0).ScaleWidth
    hdcHeight = Me.imgEating(0).ScaleHeight
    hdcWidth2 = Me.imgEating(0).ScaleWidth
    hdcHeight2 = Me.imgEating(0).ScaleHeight
    
                
    ReDim varArray(UBound(m_Cows), 1)
    For cowIndex = LBound(m_Cows) To UBound(m_Cows)
        varArray(cowIndex, 0) = m_Cows(cowIndex).YPosition
        varArray(cowIndex, 1) = cowIndex
    Next cowIndex
    ' Sort the temporary array by ascending order on the first dimension.
    ' Remember, this is a two-dimensional array.
    Call FastQSort(varArray)
    
    
    Me.Cls
    
    ' Loop through all cows in the array, from the lowest to the highest.
    For cowIndex = LBound(m_Cows) To UBound(m_Cows)
                
        m_Cows(varArray(cowIndex, 1)).StateIndex = m_Cows(varArray(cowIndex, 1)).StateIndex + 1
        
        Select Case m_Cows(varArray(cowIndex, 1)).State
                        
            Case CowState.CowWalking
                m_Cows(varArray(cowIndex, 1)).XPosition = m_Cows(varArray(cowIndex, 1)).XPosition + m_Cows(varArray(cowIndex, 1)).CowSpeed
                If m_Cows(varArray(cowIndex, 1)).XPosition > Me.ScaleWidth Then
                    m_Cows(varArray(cowIndex, 1)).XPosition = -Me.imgWalking(0).ScaleWidth
                    m_Cows(varArray(cowIndex, 1)).YPosition = Rnd * Me.ScaleHeight
                    m_Cows(varArray(cowIndex, 1)).Enabled = True
                End If
                If m_Cows(varArray(cowIndex, 1)).StateIndex > 7 Then
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 0
                End If
                
                If m_Cows(varArray(cowIndex, 1)).Enabled = True Then Call TransparentBlt(Me.hDC, m_Cows(varArray(cowIndex, 1)).XPosition, m_Cows(varArray(cowIndex, 1)).YPosition, hdcWidth2, hdcHeight2, Me.imgWalking(m_Cows(varArray(cowIndex, 1)).StateIndex).hDC, 0, 0, hdcWidth2, hdcHeight2, RGB(255, 0, 255))
                If Rnd > 0.98 Then
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 0
                    m_Cows(varArray(cowIndex, 1)).State = CowState.CowMooing
                ElseIf Rnd > 0.98 Then
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 0
                    m_Cows(varArray(cowIndex, 1)).State = CowState.CowEatingStart
                End If
                
            Case CowState.CowMooing
                If m_Cows(varArray(cowIndex, 1)).StateIndex > 10 Then
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 0
                    m_Cows(varArray(cowIndex, 1)).State = Int(Rnd * 3)
                End If
                If m_Cows(varArray(cowIndex, 1)).Enabled = True Then Call TransparentBlt(Me.hDC, m_Cows(varArray(cowIndex, 1)).XPosition, m_Cows(varArray(cowIndex, 1)).YPosition, hdcWidth2, hdcHeight2, Me.imgMooing(m_Cows(varArray(cowIndex, 1)).StateIndex).hDC, 0, 0, hdcWidth2, hdcHeight2, RGB(255, 0, 255))
                
                
            Case CowState.CowEatingStart
                If m_Cows(varArray(cowIndex, 1)).StateIndex > 4 Then
                    m_Cows(varArray(cowIndex, 1)).State = CowEating
                End If
                If m_Cows(varArray(cowIndex, 1)).Enabled = True Then Call TransparentBlt(Me.hDC, m_Cows(varArray(cowIndex, 1)).XPosition, m_Cows(varArray(cowIndex, 1)).YPosition, hdcWidth2, hdcHeight2, Me.imgEating(m_Cows(varArray(cowIndex, 1)).StateIndex).hDC, 0, 0, hdcWidth2, hdcHeight2, RGB(255, 0, 255))
            

            Case CowState.CowEating
                intRND = Int(Rnd * 2) + 4
                If m_Cows(varArray(cowIndex, 1)).Enabled = True Then Call TransparentBlt(Me.hDC, m_Cows(varArray(cowIndex, 1)).XPosition, m_Cows(varArray(cowIndex, 1)).YPosition, hdcWidth2, hdcHeight2, Me.imgEating(intRND).hDC, 0, 0, hdcWidth2, hdcHeight2, RGB(255, 0, 255))
                If Rnd > 0.99 Then
                    m_Cows(varArray(cowIndex, 1)).State = CowEatingFinish
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 3
                End If

            Case CowState.CowEatingFinish
                If m_Cows(varArray(cowIndex, 1)).StateIndex > 8 Then
                    m_Cows(varArray(cowIndex, 1)).StateIndex = 0
                    m_Cows(varArray(cowIndex, 1)).State = CowWalking
                End If
                If m_Cows(varArray(cowIndex, 1)).Enabled = True Then Call TransparentBlt(Me.hDC, m_Cows(varArray(cowIndex, 1)).XPosition, m_Cows(varArray(cowIndex, 1)).YPosition, hdcWidth2, hdcHeight2, Me.imgEating(m_Cows(varArray(cowIndex, 1)).StateIndex).hDC, 0, 0, hdcWidth2, hdcHeight2, RGB(255, 0, 255))
            
        End Select

    Next cowIndex

    Me.Refresh

End Sub


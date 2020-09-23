VERSION 5.00
Begin VB.UserControl ucVeryWellsStatusBarXP 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ControlContainer=   -1  'True
   PropertyPages   =   "ucVeryWellsStatusBarXP.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ToolboxBitmap   =   "ucVeryWellsStatusBarXP.ctx":003F
End
Attribute VB_Name = "ucVeryWellsStatusBarXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   ucVeryWellsStatusBarXP.ctl V. 1.1
'

'   Origin:                     xpWellsStatusBar by  Richard Wells

'   Redesign and extensions:    Light Templer   (Reach me on GMX :  schwepps_bitterlemon@gmx.de)
'
'   Addition on 26.5.2003       MICK-S makes an update to XPwellsStatusBar. He added the Office XP apperance and the
'                               UseWindowsColors property. I put it into VeryWellsStatusBarXP to have a All-In-One solution.
'                               Examples updated to show most of the feautures.
'
'   Special thanx to:           Keith 'LaVolpe' Fox for the API timer code.
'                               Steve 'vbAaccelerator' McMahon for details on icons and lots of stuff.
'                               Carlos 'mztools' Quintero for his great freeware VB addin.
'
'   Last changes by LT          1.6.2003
'

'   Historie:                   1.6.2003 * MEZ Fixed a bug in UserControl_Click()/UserControl_DblClick() handling
'                                          events on disabled panels. Thx to 'Dream' !
'                               2.6.2003 * Some improvements to draw_gradient().
'                                        * Added  "Public Sub ClearPanel(lPanelIndex As Long)" to erase the text
'                                          on a panel without a total redraw. Used it immediatly for PanelCaption.
'                                          This shortens the time for the API timer event very much!
'                               3.6.2003 * Three new panel types (to be 'complete' ;)) :  NUMLOCK, SCROLL, CAPSLOCK
'                                          (Does anybody really use this ?)
'                                        * Tags for panels included.
'
'                               3.6.2003 * Changed the Read/Write property strategy. This speeds things up and should solve
'                                          the problems some people have.
'                                        * Two brand new styles (Appearance): XP Diagonal Left + Right ! A tribute to
'                                          the 'LaVolpe-Button' Keith Fox wrote. Now you can use the diagonal styled buttons
'                                          on a diagonal styled statusbar!
'                                        * Changed revision to 1.1
'
'                               4.6.2003 * Adds to panels: Visible and MinWidth property. DemoForm changed to show this.
'
'


Option Explicit

' *******************************
' *            EVENTS           *
' *******************************
Public Event MouseDownInPanel(iPanel As Long)
Public Event Click(iPanelNumber)
Public Event DblClick(iPanelNumber)
Public Event TimerBeforeRedraw()
Public Event TimerAfterRedraw()
Public Event BeforeRedraw()
Public Event AfterRedraw()


' *************************************
' *        PUBLIC ENUMS               *
' *************************************
Public Enum enVWsbXPApperance           ' "Apperance" is a too common name for a public var, so I added some unique stuff.
    [Office XP] = 0
    [Windows XP] = 1
    [Simple] = 2
    [XP Diagonal Left] = 3
    [XP Diagonal Right] = 4
End Enum



' *************************************
' *   PRIVATE CONSTS (DEFAULTS)       *
' *************************************
Private Const m_def_UseWindowsColors = False
Private Const m_def_Apperance = [Windows XP]


' *************************************
' *         PRIVATE TYPE              *
' *************************************

Private Type BITMAPINFOHEADER '40 bytes
        biSize          As Long
        biWidth         As Long
        biHeight        As Long
        biPlanes        As Integer
        biBitCount      As Integer
        biCompression   As Long
        biSizeImage     As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed       As Long
        biClrImportant  As Long
End Type

Private Type RGBQUAD
        rgbBlue         As Byte
        rgbGreen        As Byte
        rgbRed          As Byte
        rgbReserved     As Byte
End Type

Private Type BITMAPINFO
        bmiHeader       As BITMAPINFOHEADER
        bmiColors(1)    As RGBQUAD
End Type


' *************************************
' *        API DEFINITIONS            *
' *************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function CreateRectRgnIndirect Lib "gdi32" _
        (lpRect As RECT) As Long

Private Declare Function PtInRegion Lib "gdi32" _
        (ByVal hRgn As Long, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function OffsetRect Lib "user32" _
        (lpRect As RECT, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function CopyRect Lib "user32" _
        (lpDestRect As RECT, _
         lpSourceRect As RECT) As Long

Private Declare Function StretchBlt Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal hSrcDC As Long, _
         ByVal xSrc As Long, _
         ByVal ySrc As Long, _
         ByVal nSrcWidth As Long, _
         ByVal nSrcHeight As Long, _
         ByVal dwRop As Long) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
        (ByVal hwnd As Long, _
         ByVal lpString As String, _
         ByVal hData As Long) As Long

Private Declare Function SetTimer Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal nIDEvent As Long, _
         ByVal uElapse As Long, _
         ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal nIDEvent As Long) As Long
         
Private Declare Function DrawEdge Lib "user32" _
        (ByVal hDC As Long, _
         qrc As RECT, _
         ByVal edge As Long, _
         ByVal grfFlags As Long) As Long

Private Declare Function InflateRect Lib "user32" _
        (lpRect As RECT, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long

Private Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal hSrcDC As Long, _
         ByVal xSrc As Long, _
         ByVal ySrc As Long, _
         ByVal dwRop As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" _
        (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" _
        (ByVal hwnd As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" _
        (ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal nPlanes As Long, _
         ByVal nBitCount As Long, _
         lpBits As Any) As Long

Private Declare Function GetBkColor Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function GetTextColor Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function SelectPalette Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal hPalette As Long, _
         ByVal bForceBackground As Long) As Long

Private Declare Function RealizePalette Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal hDC As Long) As Long

Private Declare Function CreateHalftonePalette Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" _
        (ByVal hDC As Long, _
         pBitmapInfo As BITMAPINFO, _
         ByVal un As Long, _
         ByVal lplpVoid As Long, _
         ByVal handle As Long, _
         ByVal dw As Long) As Long

Private Declare Function SetDIBColorTable Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal un1 As Long, _
         ByVal un2 As Long, _
         pcRGBQuad As RGBQUAD) As Long

Private Declare Function SetMapMode Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal nMapMode As Long) As Long

Private Declare Function GetMapMode Lib "gdi32" _
        (ByVal hDC As Long) As Long

Private Declare Function GetClientRect Lib "user32" _
        (ByVal hwnd As Long, _
         lpRect As RECT) As Long

Private Declare Function SetTextColor Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal crColor As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hDC As Long, _
         ByVal lpStr As String, _
         ByVal nCount As Long, _
         lpRect As RECT, _
         ByVal wFormat As Long) As Long

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal crColor As Long) As Long

Private Declare Function GetKeyboardState Lib "user32" _
        (kbArray As KeyboardBytes) As Long


' *************************************
' *        API CONSTANTS              *
' *************************************

' For DrawText
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_WORD_ELLIPSIS = &H40000


' Win32 edge draw consts
Private Const BF_BOTTOM = &H8
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_TOP = &H2
Private Const BF_LEFT = &H1
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)

' Win32 Special color values
Private Const CLR_INVALID = -1
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

' DIB Section constants
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

' Raster Operation Codes
Private Const DSna = &H220326 '0x00220326
Private Const ScrCopy = &HCC0020

' VB Errors
Private Const giINVALID_PICTURE As Integer = 481

' Misc
Private Const VK_CAPITAL = &H14
Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91
Private Const vbGray = 8421504

' *************************************
' *            PRIVATES               *
' *************************************

Private Enum DrawTextFlags
    [Word Break] = DT_WORDBREAK
    [Center] = DT_CENTER
    [Use Ellipsis] = DT_WORD_ELLIPSIS
End Enum

Private Type KeyboardBytes
    kbByte(0 To 255)            As Byte
End Type
Private kbArray As KeyboardBytes


' Gripper Stuff
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17

Private bDrawGripper            As Boolean
Private frm                     As Form
Private WithEvents eForm        As Form
Attribute eForm.VB_VarHelpID = -1
Private rcGripper               As RECT
Private bDrawSeperators         As Boolean
Private m_TopLine               As Boolean

' Panel Stuff.
Private m_Panels()              As New clsPanels
Private m_PanelCount            As Long
Private rcPanel()               As RECT
    
' Used for Click and DblClick Events
Private PanelNum                As Long

' Panel colors and global mask color.
Private oBackColor              As OLE_COLOR
Private oForeColor              As OLE_COLOR
Private oMaskColor              As OLE_COLOR
Private oDissColor              As OLE_COLOR


' Misc stuff
Private flgTimerEnabled         As Boolean
Private m_UseWindowsColors      As Boolean
Private m_Apperance             As enVWsbXPApperance

Private m_hpalHalftone          As Long
Private flgRedrawEnabled      As Boolean              ' Set to FALSE to prevent redawing the statusbar, don't forget
                                                        ' to re-activate! Used in 'Usercontrol_ReadProperties(), ...
'
'
'




' *************************************
' *            INIT/TERM              *
' *************************************

Private Sub UserControl_Initialize()

    flgRedrawEnabled = True
    
End Sub



Private Sub UserControl_Terminate()
        
    Dim i As Long
    
    flgRedrawEnabled = False
    
    ' Stop timer
    KillTimer UserControl.hwnd, 2201
    flgTimerEnabled = False
        
    Set frm = Nothing
    Erase rcPanel
    
End Sub





' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************

Public Sub DrawGripper()
    
    Dim lColorHighLite  As Long
    Dim lColorShaddow   As Long
    Dim lColorGrad      As Long
    Dim i               As Long
    
    With rcGripper
        .lLeft = UserControl.ScaleWidth - 15
        .lRight = UserControl.ScaleWidth
        .lBottom = UserControl.ScaleHeight
        .lTop = UserControl.ScaleHeight - 15
    End With
    
    With UserControl
                
        ' HiLite and Shaddow color
        If m_UseWindowsColors = True Then
            lColorHighLite = TranslateColorToRGB(GetSysColor(COLOR_BTNHIGHLIGHT), 0, 0, 0)
            lColorShaddow = TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
        Else
            lColorHighLite = TranslateColorToRGB(.BackColor, 0, 0, 0, -50)
            lColorShaddow = TranslateColorToRGB(.BackColor, 0, 0, 0, 50)
        End If
        
        Select Case m_Apperance
            
            Case [Windows XP]
                    ' Retain the area
                    DrawASquare .hDC, rcGripper, .BackColor, True
                    
                    DrawALine .hDC, rcGripper.lLeft, rcGripper.lBottom - 1, rcGripper.lRight, rcGripper.lBottom - 1, _
                            TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2

                    DrawALine .hDC, rcGripper.lLeft, rcGripper.lBottom - 3, rcGripper.lRight, rcGripper.lBottom - 3, _
                            TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2
                    
                    DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 3, .ScaleWidth - 3, .ScaleHeight - 3, lColorShaddow, 2
                    DrawALine .hDC, .ScaleWidth - 7, .ScaleHeight - 3, .ScaleWidth - 7, .ScaleHeight - 3, lColorShaddow, 2
                    DrawALine .hDC, .ScaleWidth - 11, .ScaleHeight - 3, .ScaleWidth - 11, .ScaleHeight - 3, lColorShaddow, 2
                
                    DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 7, .ScaleWidth - 3, .ScaleHeight - 7, lColorShaddow, 2
                    DrawALine .hDC, .ScaleWidth - 7, .ScaleHeight - 7, .ScaleWidth - 7, .ScaleHeight - 7, lColorShaddow, 2
                
                    DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 11, .ScaleWidth - 3, .ScaleHeight - 11, lColorShaddow, 2
                
                    DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 4, .ScaleWidth - 4, .ScaleHeight - 4, lColorHighLite, 2
                    DrawALine .hDC, .ScaleWidth - 8, .ScaleHeight - 4, .ScaleWidth - 8, .ScaleHeight - 4, lColorHighLite, 2
                    DrawALine .hDC, .ScaleWidth - 12, .ScaleHeight - 4, .ScaleWidth - 12, .ScaleHeight - 4, lColorHighLite, 2
                
                    DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 8, .ScaleWidth - 4, .ScaleHeight - 8, lColorHighLite, 2
                    DrawALine .hDC, .ScaleWidth - 8, .ScaleHeight - 8, .ScaleWidth - 8, .ScaleHeight - 8, lColorHighLite, 2
                
                    DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 12, .ScaleWidth - 4, .ScaleHeight - 12, lColorHighLite, 2
        
        
            Case [Office XP]
                    ' Retain the area
                    DrawASquare .hDC, rcGripper, .BackColor, True
                    
                    For i = 5 To 15 Step 5
                        DrawALine .hDC, .ScaleWidth - i, .ScaleHeight, .ScaleWidth, .ScaleHeight - i, lColorHighLite
                    Next i
                    
                    For i = 2 To 14
                        If i = 5 Or i = 10 Then
                            i = i + 2
                        End If
                        DrawALine .hDC, .ScaleWidth - i, .ScaleHeight, .ScaleWidth, .ScaleHeight - i, lColorShaddow
                    Next i
                                                      
            
            Case [Simple]
                    ' In progress ... ;)
                    For i = 2 To 14
                        DrawALine .hDC, .ScaleWidth - i, .ScaleHeight, .ScaleWidth, .ScaleHeight - i, oForeColor
                    Next i
            
            
            Case [XP Diagonal Left], [XP Diagonal Right]
                    For i = 3 To 13
                        lColorGrad = 140 + (6 * i)
                        DrawALine .hDC, .ScaleWidth - i, .ScaleHeight - 3, .ScaleWidth - 1, .ScaleHeight - i, _
                                RGB(lColorGrad, lColorGrad, lColorGrad)
                    Next i

        End Select
        
        UserControl.Refresh
    End With

End Sub

Public Function InsertPanel(ByVal lCurrentPanel As Long) As Long

    Dim i As Long

    m_PanelCount = m_PanelCount + 1
    ReDim Preserve m_Panels(1 To m_PanelCount) As New clsPanels
    
    ' Make space for the new one
    lCurrentPanel = lCurrentPanel + 1
    For i = m_PanelCount To lCurrentPanel + 1 Step -1
        Set m_Panels(i) = m_Panels(i - 1)
    Next i
    Set m_Panels(lCurrentPanel) = New clsPanels
    With m_Panels(lCurrentPanel)
        .ClientWidth = 100
        .pEnabled = True
        Set .PanelPicture = Nothing
        .PanelEdgeInner = 0
        .PanelEdgeOuter = 0
        .PanelEdgeSpacing = 0
        .PanelGradient = 0
    End With
    PropertyChanged "NumberOfPanels"
    InsertPanel = m_PanelCount
    DrawStatusBar True
    
End Function

Public Function DeletePanel(lPanelIndex As Long)
    
    Dim i As Long
    
    If m_PanelCount > 0 Then
        For i = lPanelIndex To m_PanelCount - 1
            Set m_Panels(i) = m_Panels(i + 1)
        Next i
        Set m_Panels(m_PanelCount) = Nothing
        m_PanelCount = m_PanelCount - 1
        If m_PanelCount > 0 Then
            ReDim Preserve m_Panels(1 To m_PanelCount)
        Else
            Erase m_Panels()
        End If
        PropertyChanged "NumberOfPanels"
        DrawStatusBar True
    End If
    
End Function

Public Sub DrawStatusBar(Optional FullRedraw As Boolean = True)

    Dim i               As Long
    Dim rc              As RECT
    Dim rcTemp          As RECT
    Dim X               As Long
    Dim Y               As Long
    Dim X1              As Long
    Dim Y1              As Long
    Dim lOffset         As Long
    Dim pX              As Long
    Dim pY              As Long
    Dim lColorTmp1      As Long
    Dim lColorTmp2      As Long
    Dim lSpringer       As Long
    Dim lFixedSizeTotal As Long
    Dim lSpringSize     As Long
    Dim lPPxPos         As Long
    Dim ContainedCtrl   As Control
    Dim lGapToBorder    As Long         ' Controls distance to top/bottom for panel fillings (gradients ...)
    Dim ltmp            As Long

    
    On Error GoTo error_handler
    
    If flgRedrawEnabled = False Then  ' Prevent redrawing during lot of property changes like in 'Usercontrol_ReadProperties()'
        
        Exit Sub
    End If

    RaiseEvent BeforeRedraw
    
    
    If FullRedraw = True Then
        With UserControl
            ' == Control Shading Lines ==
            Cls
            
            Select Case m_Apperance

                Case [Office XP]

                Case [Windows XP]
                        ' Top lines
                        If m_TopLine = True Then
                            DrawALine .hDC, 0, 0, .ScaleWidth, 0, TranslateColorToRGB(oBackColor, 0, 0, 0, -45)
                        End If
                        lOffset = 36
                        For i = 1 To 4
                            DrawALine .hDC, 0, i, .ScaleWidth, i, TranslateColorToRGB(oBackColor, 0, 0, 0, lOffset)
                            lOffset = lOffset - 9
                        Next i
                                    
                        ' Bottom Lines
                        DrawALine .hDC, 0, .ScaleHeight - 1, .ScaleWidth, .ScaleHeight - 1, _
                                TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2
                        DrawALine .hDC, 0, .ScaleHeight - 3, .ScaleWidth, .ScaleHeight - 3, _
                                TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2

                        
                Case [Simple]
                        If m_TopLine = True Then
                            DrawALine .hDC, 0, 0, .ScaleWidth, 0, vbBlack
                        End If
                
                
                Case [XP Diagonal Left], [XP Diagonal Right]
                
                        lColorTmp1 = RGB(90, 90, 90)
                        
                        ' Top lines
                        DrawALine .hDC, 2, 0, .ScaleWidth - 2, 0, lColorTmp1
                        DrawALine .hDC, 2, 1, .ScaleWidth - 1, 1, vbWhite
                        DrawALine .hDC, 2, 2, .ScaleWidth - 1, 2, RGB(248, 248, 248)
                        
                        DrawVertGradient RGB(240, 240, 240), RGB(220, 220, 220), 1, .ScaleWidth - 2, 3, .ScaleHeight - 3
                        
                        ' Bottom Lines
                        DrawALine .hDC, 2, .ScaleHeight - 3, .ScaleWidth - 1, .ScaleHeight - 3, RGB(217, 217, 217)
                        DrawALine .hDC, 2, .ScaleHeight - 2, .ScaleWidth - 1, .ScaleHeight - 2, RGB(190, 190, 190)
                        DrawALine .hDC, 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, lColorTmp1
                
                        ' Left lines
                        DrawALine .hDC, 0, 2, 0, .ScaleHeight - 2, lColorTmp1
                        DrawALine .hDC, 1, 2, 1, .ScaleHeight - 2, RGB(230, 230, 230)
                        
                        ' Right lines
                        DrawALine .hDC, .ScaleWidth - 2, 2, .ScaleWidth - 2, .ScaleHeight - 2, RGB(230, 230, 230)
                        DrawALine .hDC, .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, lColorTmp1
                        
                        ' Draw dots into corners
                        SetPixel .hDC, 1, 1, lColorTmp1
                        SetPixel .hDC, .ScaleWidth - 2, 1, lColorTmp1
                        SetPixel .hDC, 1, .ScaleHeight - 2, lColorTmp1
                        SetPixel .hDC, .ScaleWidth - 2, .ScaleHeight - 2, lColorTmp1
                        
            End Select
            
            
        End With
    End If
    
    ' The Panels
    '******************* Dimensions. **********************
    ' X = Left of the panel
    ' Y = Top of the panel
    ' X1 = Width of the panel
    ' Y1 = Height of the panel
    '******************************************************
    
    Select Case m_Apperance

        Case [Office XP]
                Y = 1                               ' Start the panel 1 pixel down from the top edge.
                Y1 = UserControl.ScaleHeight - 1    ' Height of the panel
    
                
        Case [Windows XP]
                Y = 5                               ' Start the panel 5 pixels down from the top edge.
                Y1 = UserControl.ScaleHeight - 4    ' Height of the panel
    
        Case [Simple], [XP Diagonal Left], [XP Diagonal Right]
                Y = 1                               ' Start the panel 1 pixel down from the top edge.
                Y1 = UserControl.ScaleHeight - 1    ' Height of the panel
                
    End Select
    
    
    ' Two tasks for this loop:
    '               1 - How many panels with PanelType = [PT Text spring size] we have ?
    '               2 - Adjust panels size with PanelType = [PT Text AutoSize contents]
    lSpringer = 0
    For i = 1 To m_PanelCount
        With m_Panels(i)
            If .pVisible = True Then
                Select Case .PanelType
                
                    Case [PT Text spring size]
                            lSpringer = lSpringer + 1
                        
                        
                    Case [PT Text AutoSize contents]
                            .ClientWidth = UserControl.TextWidth(.PanelText) + (UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode)) + 12
                            lFixedSizeTotal = lFixedSizeTotal + .ClientWidth     ' Get total size of fixed-size panels
                            
                            
                    Case Else
                            lFixedSizeTotal = lFixedSizeTotal + .ClientWidth     ' Get total size of fixed-size panels
                            
                End Select
            End If
        End With
    Next i
    
    ' If we have spring panels:  Adjust the width of all! spring panels
    If lSpringer > 0 Then
    
        lSpringSize = (UserControl.ScaleWidth - (lFixedSizeTotal + IIf(bDrawGripper = True, 17, 5))) / lSpringer
        If lSpringSize < 0 Then
            lSpringSize = 0
        End If
        
        For i = 1 To m_PanelCount
            With m_Panels(i)
                If .PanelType = [PT Text spring size] And .pVisible = True Then
                    .ClientWidth = IIf(lSpringSize > .pMinWidth, lSpringSize, .pMinWidth)
                End If
            End With
        Next i
        
    End If
    
    
    ' Loop through the panels
    For i = 1 To m_PanelCount
        With m_Panels(i)
        
            If .pVisible = True Then
        
                'Position the panel.
                .ClientLeft = X
                .ClientTop = Y
                
                X1 = .ClientWidth
                .ClientHeight = Y1
                
                            
                'Create a RECT area using the above dimentions to draw into.
                With rc
                    .lLeft = X
                    .lTop = Y
                    .lRight = .lLeft + X1
                    .lBottom = Y1
                End With
                ReDim Preserve rcPanel(i)
                rcPanel(i) = rc
                InflateRect rcPanel(i), -2, 0
            
                With UserControl
                    If FullRedraw = True And bDrawSeperators = True Then
                    
                        Select Case m_Apperance
        
                            Case [Office XP]
                                    
                                    
                            Case [Windows XP]
                                lColorTmp1 = TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
                                lColorTmp2 = TranslateColorToRGB(oBackColor, 0, 0, 0, -50)
                            
                                ' Draw the seperators taking into acount the first and last
                                ' panel seperators are different.
                                If i <> 1 Then
                                    ' This will draw the left line ( The lighter shade )
                                    ' so the first panel does not need one
                                    DrawALine .hDC, X, Y, X, Y1, lColorTmp1
                                End If
                                
                                If i <> m_PanelCount Then
                                    ' This will draw the right line ( The darker shade )
                                    ' Every panel will have this line exept the last
                                    ' panel has this line positioned differently.
                                    DrawALine .hDC, rc.lRight - 1, Y, rc.lRight - 1, Y1, lColorTmp2
                                Else
                                    ' Lines for the last panel.
                                    DrawALine .hDC, rc.lRight - 1, Y, rc.lRight - 1, Y1, lColorTmp1
                                    DrawALine .hDC, rc.lRight - 2, Y, rc.lRight - 2, Y1, lColorTmp2
                                End If
                    
                            
                        Case [Simple]
                                DrawALine .hDC, X, Y, X, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
                                
                                
                        Case [XP Diagonal Left]
                                If i > 1 Then
                                    lOffset = (Y1 / 2 = Int(Y1 / 2))    ' even or odd ?
                                    ltmp = Y1 \ 2
                                    DrawALine .hDC, X - ltmp - 1, Y, X + ltmp + lOffset - 1, Y1, vbGray
                                    DrawALine .hDC, X - ltmp, Y, X + ltmp + lOffset, Y1, RGB(90, 90, 90)
                                    DrawALine .hDC, X - ltmp + 1, Y, X + ltmp + lOffset + 1, Y1, vbWhite
                                    UserControl.Refresh
                                End If
            
                        Case [XP Diagonal Right]
                                If i > 1 Then
                                    lOffset = (Y1 / 2 = Int(Y1 / 2))    ' even or odd ?
                                    ltmp = Y1 \ 2
                                    DrawALine .hDC, X + ltmp - 1, Y, X - ltmp - lOffset - 1, Y1, vbGray
                                    DrawALine .hDC, X + ltmp, Y, X - ltmp - lOffset, Y1, RGB(90, 90, 90)
                                    DrawALine .hDC, X + ltmp + 1, Y, X - ltmp - lOffset + 1, Y1, vbWhite
                                    UserControl.Refresh
                                End If
                                
                        End Select
                        
                    End If
                    
                
                    ' Design the panel
                    Select Case m_Apperance
        
                        Case [Office XP]
                                ' DrawASquare UserControl.hDC, rcPanel(i), oBackColor, True  (Not sure when needed ... LT)
                                DrawASquare .hDC, rcPanel(i), vbButtonShadow, False
                                
                                            
                        Case [Windows XP]
                                ' DrawASquare .hDC, rcPanel(i), oBackColor, True  (Not sure when needed ... LT)
                                
                        Case [Simple]
                                X = X + 2
        
                    End Select
                End With
                
                ' ### Maybe we want to draw some fancy background gradients and framing stuff ;) ... ###
                InflateRect rc, -3, -2
                            
                ' Gradients ?
                lGapToBorder = UserControl.ScaleHeight / 7
                Select Case .PanelGradient
    
                        Case 1      ' [Transparent]     :  So do nothing ;)
                
                
                        Case 2      ' [Opaque]          :  Draw a simple rectangle in panel background color
                                    CopyRect rcTemp, rc
                                    rcTemp.lLeft = rcTemp.lLeft + 1
                                    rcTemp.lRight = rcTemp.lRight - 2
                                    DrawASquare UserControl.hDC, rcTemp, .PanelBckgColor, True
                                    
                
                        Case 3      ' [Top Bottom]      :  Simple gradient 1
                                    DrawVertGradient .PanelBckgColor, vbWhite, _
                                            X + 3, .ClientWidth - 7, _
                                            lGapToBorder, UserControl.ScaleHeight - lGapToBorder
                                            
                                            
                        Case 4      ' [Top 1/3 Bottom]  :  Complex gradient 1
                                    DrawVertGradient .PanelBckgColor, vbWhite, _
                                            X + 3, .ClientWidth - 7, _
                                            lGapToBorder, UserControl.ScaleHeight / 3 + 2
                                            
                                    DrawVertGradient vbWhite, .PanelBckgColor, _
                                            X + 3, .ClientWidth - 7, _
                                            UserControl.ScaleHeight / 3 + 2, UserControl.ScaleHeight - lGapToBorder
                                            
                                            
                        Case 5      ' [Top 1/2 Bottom]  :  Complex gradient 2
                                    DrawVertGradient .PanelBckgColor, vbWhite, _
                                            X + 3, .ClientWidth - 7, _
                                            lGapToBorder, UserControl.ScaleHeight / 2
                                            
                                    DrawVertGradient vbWhite, .PanelBckgColor, _
                                            X + 3, .ClientWidth - 7, _
                                            UserControl.ScaleHeight / 2, UserControl.ScaleHeight - lGapToBorder
                                                
                                                
                        Case 6      ' [Top 2/3 Bottom]  :  Complex gradient 3
                                    DrawVertGradient .PanelBckgColor, vbWhite, _
                                            X + 3, .ClientWidth - 7, _
                                            lGapToBorder, (UserControl.ScaleHeight / 3) * 2 - 3
                                            
                                    DrawVertGradient vbWhite, .PanelBckgColor, _
                                            X + 3, .ClientWidth - 7, _
                                            (UserControl.ScaleHeight / 3) * 2 - 2, UserControl.ScaleHeight - lGapToBorder
                                            
                                            
                        Case 7      ' [Bottom Top]      :  Simple gradient 2
                                    DrawVertGradient vbWhite, .PanelBckgColor, _
                                            X + 3, .ClientWidth - 7, _
                                            lGapToBorder, UserControl.ScaleHeight - lGapToBorder
                                                
                                                
                        Case Else   ' Just for sure ;)
                        
                                    
                End Select
    
            
                ' Draw the OUTER Edge
                rc.lTop = lGapToBorder
                rc.lBottom = UserControl.ScaleHeight - (lGapToBorder - 2)
                DrawEdge UserControl.hDC, rc, .PanelEdgeOuter, BF_TOPLEFT
                DrawEdge UserControl.hDC, rc, .PanelEdgeOuter, BF_BOTTOMRIGHT
                
                ' make rectangle smaller by inner spacing property
                InflateRect rc, -.PanelEdgeSpacing, -.PanelEdgeSpacing
                
                ' Draw the INNER Edge
                DrawEdge UserControl.hDC, rc, .PanelEdgeInner, BF_TOPLEFT
                DrawEdge UserControl.hDC, rc, .PanelEdgeInner, BF_BOTTOMRIGHT
                            
                
                
                GetPanelPictureSize i, pX, pY   ' Get the size of the picture even if there is no one set
        
                ' Create a temporary RECT to draw some text into.
                GetClientRect UserControl.hwnd, rcTemp
                DrawText UserControl.hDC, .PanelText, Len(.PanelText), rcTemp, DT_CALCRECT Or DT_WORDBREAK
                CopyRect rc, rcTemp
                
                ' Position our RECT
                Select Case .PanelPicAlignment
                
                    Case [PP Left]
                            rc.lLeft = X + pX + 2
                            rc.lRight = ((rc.lLeft + X1) - 10) - pX
                                
                    Case [PP Center]
                            rc.lLeft = X
                            rc.lRight = ((rc.lLeft + X1) - 10)
                            
                    Case [PP Right]
                            rc.lLeft = X
                            rc.lRight = ((rc.lLeft + X1) - 10) - pX
                            
                End Select
                
                If .PanelEdgeOuter <> 0 Then
                    InflateRect rc, -3, 0
                End If
                If .PanelEdgeInner <> 0 Then
                    InflateRect rc, -(.PanelEdgeSpacing + 3), 0
                End If
                
                ' Save this contents area !
                .ContentsLeft = rc.lLeft
                .ContentsTop = rc.lTop
                .ContentsRight = rc.lRight
                .ContentsBottom = rc.lBottom
                
                
                ' Draw the text into our new panel.
                SetTextColor UserControl.hDC, IIf(.pEnabled = True, oForeColor, oDissColor)
                OffsetRect rc, 4, (ScaleHeight - rc.lBottom) / 2
                DrawTheText UserControl.hDC, .PanelText, Len(.PanelText), rc, .TextAlignment
                
            
                ' Add a PanelPicture if required.
                
                ' TODO :
                '           Picture will spill into the next panel if for some reason someone
                '           sets the PanelWidth to a smaller width than the image.
                '
                '           Seems not really be a great prob... ;) - LT
                
                If Not (.PanelPicture Is Nothing) Then
                
                    lPPxPos = Choose(.PanelPicAlignment + 1, _
                            IIf(.PanelEdgeInner = 0, X + 5, X + 7 + .PanelEdgeSpacing), _
                            X + (.ClientWidth / 2) - (pX / 2), _
                            (X + .ClientWidth) - (pX + 5 + IIf(.PanelEdgeInner = 0, 0, .PanelEdgeSpacing)))
                            
                    PaintTransparentPicture UserControl.hDC, _
                            .PanelPicture, _
                            lPPxPos, (ScaleHeight - pY) / 2, _
                            pX, pY, _
                            0, 0, _
                            oMaskColor
                           
                    Refresh
                            
                End If
                            
                'Dont forget to move the X for the next panel ...
                X = X + .ClientWidth
                
            End If  ' .pVisible = True
        End With
    Next i

    ' If there are integrated controls: Set position(s)             ' Magic number format "### 03 0050 +"
    On Error Resume Next                                            ' Means: Put control to panel 3, 50 twips from left panel
    For Each ContainedCtrl In UserControl.ContainedControls         ' border and adjust size in horicontaldirection. Use "-"
        With ContainedCtrl                                          ' for no adjustment e.g. "### 02 0050 -"
            If Len(.Tag) = 13 And Left$(.Tag, 4) = "### " Then      ' Handle controls with "magic number tag" only!
                i = Val(Mid$(.Tag, 5, 2))                           ' Get panel index
                
                If i > 0 And i <= m_PanelCount Then                 ' Only if we HAVE panels!
                    If m_Panels(i).pVisible = True Then
                        .Visible = True
                        X = Val(Mid$(.Tag, 8, 4))
                        .Left = UserControl.ScaleX(m_Panels(i).ContentsLeft + 3, vbPixels, vbTwips) + X
                        If Right$(.Tag, 1) = "+" Then
                            ltmp = (UserControl.ScaleX(m_Panels(i).ContentsRight, vbPixels, vbTwips)) - .Left
                            If ltmp > 0 Then
                                .Width = ltmp
                            End If
                        End If
                    Else
                        .Visible = False                            ' Don't show integrated controls on invisible panels!
                    End If
                End If
            End If
            
        End With
    Next ContainedCtrl
    
    On Error GoTo 0

    If bDrawGripper = True Then
        DrawGripper
    End If
    
    RaiseEvent AfterRedraw

    On Error GoTo 0

    Exit Sub


error_handler:

    MsgBox "Error [" + Err.Description + "] in procedure 'DrawStatusBar()' at Benutzersteuerelement ucVeryWellsStatusBarXP"
    
End Sub


Public Sub RefreshAll()
    ' Redraw all

    DrawStatusBar True

End Sub


Public Sub ClearPanel(lPanelIndex As Long)
    ' Removes the text from a panel without a complete redraw of the control (speed! ...)
    ' This is done by copying the pixel colume left to the text to the whole area using
    ' the StretchBlt() API function

    Dim lSrcX   As Long
    Dim lWidth  As Long
    Dim lHeight As Long

    If lPanelIndex < 1 Or lPanelIndex > m_PanelCount Then
        
        Exit Sub
    End If
    
    With m_Panels(lPanelIndex)
        lSrcX = .ContentsLeft
        lWidth = .ContentsRight - lSrcX
        lHeight = .ClientHeight
    End With
    
    StretchBlt UserControl.hDC, lSrcX + 1, 0, lWidth + 1, lHeight, UserControl.hDC, lSrcX, 0, 1, lHeight, ScrCopy
    UserControl.Refresh
    
End Sub


' *************************************
' *         FRIEND FUNCTIONS          *
' *************************************

Friend Sub TimerUpdate()

    Dim i   As Long
    Dim rc  As RECT
    
    RaiseEvent TimerBeforeRedraw
    
    For i = 1 To m_PanelCount
        With m_Panels(i)
            Select Case .PanelType
            
            Case [PT Date]
                    PanelCaption(i) = Format(Date, "d.m.yyyy")
                                            
            Case [PT Time]
                    PanelCaption(i) = Format(Time, "hh:nn:ss")
            
            Case [PT CapsLock]
                    GetKeyboardState kbArray
                    .pEnabled = IIf(kbArray.kbByte(VK_CAPITAL) = 1, True, False)
                    PanelCaption(i) = "CAPS"
            
            Case [PT NumLock]
                    GetKeyboardState kbArray
                    .pEnabled = IIf(kbArray.kbByte(VK_NUMLOCK) = 1, True, False)
                    PanelCaption(i) = "NUM"
            
            Case [PT Scroll]
                    GetKeyboardState kbArray
                    .pEnabled = IIf(kbArray.kbByte(VK_SCROLL) = 1, True, False)
                    PanelCaption(i) = "SCROLL"
            
            End Select
            
        End With
    Next i
    
    RaiseEvent TimerAfterRedraw
    
End Sub



' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************

Private Sub UserControl_InitProperties()
            
    flgRedrawEnabled = False
    Set UserControl.Font = UserControl.Parent.Font
    oBackColor = vbButtonFace
    oForeColor = vbButtonText
    oDissColor = vbGrayText
    oMaskColor = RGB(255, 0, 255)
    bDrawGripper = True
    flgTimerEnabled = False
    m_Apperance = m_def_Apperance
    m_UseWindowsColors = m_def_UseWindowsColors
    flgRedrawEnabled = True
    DrawStatusBar True
     
End Sub


Private Sub UserControl_Show()

    ' Ensure "special background handling"
    Select Case m_Apperance

    Case [XP Diagonal Left], [XP Diagonal Right]
        BackColor = oBackColor = UserControl.Parent.BackColor

    End Select

    If Ambient.UserMode = True Then
        If flgTimerEnabled = False Then
            ' Start timer
            SetProp UserControl.hwnd, "sbXP_ClassID", ObjPtr(Me)
            SetTimer UserControl.hwnd, 2201, 200, AddressOf API_Timer_Callback
            flgTimerEnabled = True
        End If
    Else
        ' Stop timer
        KillTimer UserControl.hwnd, 2201
        flgTimerEnabled = False
    End If
    
End Sub


Private Sub UserControl_Click()
    
    If PanelNum < 1 Then
         
         Exit Sub
    End If
    If m_Panels(PanelNum).pEnabled = True Then
        RaiseEvent Click(PanelNum)
    End If
    
End Sub

Private Sub UserControl_DblClick()
    
    If PanelNum < 1 Then
         
         Exit Sub
    End If
    If m_Panels(PanelNum).pEnabled = True Then
        RaiseEvent DblClick(PanelNum)
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim pt      As POINTAPI
    Dim hRgn    As Long
    Dim i       As Long
    
    PanelNum = 0
    If ShowGripper = True Then
        hRgn = CreateRectRgnIndirect(rcGripper)
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            If Button = vbLeftButton Then
                SizeByGripper frm.hwnd
                DeleteObject hRgn
                
                Exit Sub
            End If
        End If
        
    End If
    
    For i = 1 To m_PanelCount
        hRgn = CreateRectRgnIndirect(rcPanel(i))
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            If Button = vbLeftButton Then
                If m_Panels(i).pEnabled = True Then
                    PanelNum = i
                    RaiseEvent MouseDownInPanel(i)
                End If
                DeleteObject hRgn
            End If
        End If
    Next i
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim hRgn    As Long
    Dim i       As Long
    
    On Error GoTo error_handler

    If ShowGripper = True Then
        hRgn = CreateRectRgnIndirect(rcGripper)
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            UserControl.MousePointer = vbSizeNWSE
            DeleteObject hRgn
            
            Exit Sub
        Else
            UserControl.MousePointer = vbArrow
            DeleteObject hRgn
        End If
    Else
        UserControl.MousePointer = vbArrow
    End If
    
    If m_PanelCount < 1 Then        ' Jut for sure ...
    
        Exit Sub
    End If
    For i = 1 To m_PanelCount
        hRgn = CreateRectRgnIndirect(rcPanel(i))
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            Extender.ToolTipText = m_Panels(i).ToolTipTxt
        End If
        DeleteObject hRgn
    Next i

    On Error GoTo 0

    Exit Sub


error_handler:
    
    ' User 'Dream' from PSC noticed an error in this sub, but I didn't get one using my VB5
    ' so to have a fast solution, I inserted this error handler. Thx for any hints!
    
    ' MsgBox "Error [" + Err.Description + "] in procedure 'UserControl_MouseMove()' at Benutzersteuerelement ucVeryWellsStatusBarXP"
    
    If hRgn Then
        DeleteObject hRgn
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim i As Long

    On Error GoTo error_handler
    
    flgRedrawEnabled = False
    With PropBag
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorDisabled = .ReadProperty("ForeColorDissabled", vbGrayText)
        MaskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
        
        Set UserControl.Font = .ReadProperty("Font", UserControl.Parent.Font)
        ShowGripper = .ReadProperty("ShowGripper", True)
        ShowSeperators = .ReadProperty("ShowSeperators", True)
        m_Apperance = .ReadProperty("Apperance", m_def_Apperance)
        m_UseWindowsColors = .ReadProperty("UseWindowsColors", m_def_UseWindowsColors)
        m_TopLine = .ReadProperty("TopLine", True)
        
        m_PanelCount = .ReadProperty("NumberOfPanels", 0)
    End With
    
    If m_PanelCount > 0 Then
        ReDim m_Panels(1 To m_PanelCount) As New clsPanels
    End If
    For i = 1 To m_PanelCount
        With m_Panels(i)
            .pEnabled = PropBag.ReadProperty("pEnabled" & i, True)
            .pVisible = PropBag.ReadProperty("pVisible" & i, True)
            .ClientWidth = PropBag.ReadProperty("PWidth" & i)
            .pMinWidth = PropBag.ReadProperty("PMinWidth" & i, 10)
            .ToolTipTxt = PropBag.ReadProperty("pTTText" & i)
            
            .PanelType = PropBag.ReadProperty("pType" & i, [PT Text spring size])
            .PanelText = PropBag.ReadProperty("pText" & i)
            .TextAlignment = PropBag.ReadProperty("pTextAlignment" & i, [TA Left])
            
            Set .PanelPicture = PropBag.ReadProperty("PanelPicture" & i, Nothing)
            .PanelPicAlignment = PropBag.ReadProperty("PanelPicAlignment" & i)
            
            .PanelBckgColor = PropBag.ReadProperty("pBckgColor" & i)
            .PanelGradient = PropBag.ReadProperty("pGradient" & i)
            .PanelEdgeSpacing = PropBag.ReadProperty("pEdgeSpacing" & i)
            .PanelEdgeInner = PropBag.ReadProperty("pEdgeInner" & i)
            .PanelEdgeOuter = PropBag.ReadProperty("pEdgeOuter" & i)
            
            .Tag = PropBag.ReadProperty("pTag" & i, vbNullString)
        End With
    Next i
    
    flgRedrawEnabled = True
    DrawStatusBar True
    
    Exit Sub
    
    
error_handler:

    If Err.Number = 327 Then        ' In Immediate Window:  err.raise 327  , then <Help> to get infos
        Err.Clear
    Else
        MsgBox "Error [" + Err.Description + "] in 'UserControl_ReadProperties()', Modul 'ucVeryWellsStatusBarXP'", _
                vbExclamation, " Fehler "
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Dim i As Long
    
    On Error GoTo error_handler
    
    flgRedrawEnabled = False
    With PropBag
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorDissabled", oDissColor
        .WriteProperty "MaskColor", oMaskColor
        
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "ShowGripper", bDrawGripper
        .WriteProperty "ShowSeperators", bDrawSeperators
        .WriteProperty "Apperance", m_Apperance, m_def_Apperance
        .WriteProperty "UseWindowsColors", m_UseWindowsColors, m_def_UseWindowsColors
        .WriteProperty "TopLine", m_TopLine, True
        
        .WriteProperty "NumberOfPanels", m_PanelCount
    End With

    For i = 1 To m_PanelCount
        With m_Panels(i)
            
            PropBag.WriteProperty "pEnabled" & i, .pEnabled, True
            PropBag.WriteProperty "pVisible" & i, .pVisible, True
            PropBag.WriteProperty "PWidth" & i, .ClientWidth
            PropBag.WriteProperty "PMinWidth" & i, .pMinWidth, 10
            PropBag.WriteProperty "pTTText" & i, .ToolTipTxt
            
            PropBag.WriteProperty "pType" & i, .PanelType
            PropBag.WriteProperty "pText" & i, .PanelText
            PropBag.WriteProperty "pTextAlignment" & i, .TextAlignment
            
            PropBag.WriteProperty "PanelPicture" & i, .PanelPicture
            PropBag.WriteProperty "PanelPicAlignment" & i, .PanelPicAlignment
            
            PropBag.WriteProperty "pBckgColor" & i, .PanelBckgColor
            PropBag.WriteProperty "pGradient" & i, .PanelGradient
            PropBag.WriteProperty "pEdgeSpacing" & i, .PanelEdgeSpacing
            PropBag.WriteProperty "pEdgeInner" & i, .PanelEdgeInner
            PropBag.WriteProperty "pEdgeOuter" & i, .PanelEdgeOuter
            
            PropBag.WriteProperty "pTag" & i, .Tag, vbNullString
                        
        End With
    Next i
    flgRedrawEnabled = True
    
    On Error GoTo 0

    Exit Sub


error_handler:

    MsgBox "Error [" + Err.Description + "] in 'UserControl_WriteProperties()', Modul 'ucVeryWellsStatusBarXP'", _
            vbExclamation, " Fehler "
    flgRedrawEnabled = True
    
End Sub

Private Sub UserControl_Resize()

    DrawStatusBar
    
End Sub

Private Sub GetPanelPictureSize(Index As Long, X As Long, Y As Long)
    
    If m_Panels(Index).PanelPicture Is Nothing Then
            
        Exit Sub
    End If
    
    With UserControl
        X = .ScaleX(m_Panels(Index).PanelPicture.Width, vbHimetric, .ScaleMode)
        Y = .ScaleY(m_Panels(Index).PanelPicture.Height, vbHimetric, .ScaleMode)
    End With
    
End Sub

Private Sub SizeByGripper(ByVal iHwnd As Long)
  
  ReleaseCapture
  SendMessage iHwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0

End Sub


Private Sub DrawTheText(DestDC As Long, sText As String, iTextLength As Long, rc As RECT, DTF As DrawTextFlags)
    
    DrawText DestDC, sText, iTextLength, rc, DTF

End Sub


Private Sub DrawVertGradient(lFromColor As Long, _
                                lToColor As Long, _
                                start_x As Long, _
                                wid As Long, _
                                start_y As Long, _
                                end_y As Long)
                                
    ' Fast draw gradient vertical lines
    
    Const PS_SOLID = 0
    
    Dim hgt             As Single
    Dim R               As Single
    Dim G               As Single
    Dim B               As Single
    Dim dR              As Single
    Dim dg              As Single
    Dim db              As Single
    Dim Y               As Single
    Dim end_r           As Single
    Dim end_g           As Single
    Dim end_b           As Single
    Dim lRight          As Long
    Dim bArray(1 To 4)  As Byte
    Dim pt              As POINTAPI
    Dim hPen            As Long
    Dim hPen1           As Long
    Dim dstDC           As Long
    
    Dim lOld            As Long
    
    lFromColor = OleToColor(lFromColor)
    CopyMemory bArray(1), lFromColor, 4
    R = bArray(1)
    G = bArray(2)
    B = bArray(3)
    
    lToColor = OleToColor(lToColor)
    CopyMemory bArray(1), lToColor, 4
    end_r = bArray(1)
    end_g = bArray(2)
    end_b = bArray(3)

    hgt = end_y - start_y
    If hgt = 0 Then
        hgt = 1
    End If
    
    dR = (end_r - R) / hgt
    dg = (end_g - G) / hgt
    db = (end_b - B) / hgt
    
    lRight = start_x + wid
    
    dstDC = UserControl.hDC
    
    With UserControl
        lOld = .ForeColor
        For Y = start_y To end_y
            .ForeColor = RGB(R, G, B)
            
            MoveToEx dstDC, start_x, Y, pt
            LineTo dstDC, lRight, Y
            
            R = R + dR
            G = G + dg
            B = B + db
    
        Next Y
        .ForeColor = lOld
    End With
    
End Sub


Private Function OleToColor(ByVal OleColor As OLE_COLOR) As Long

    If (OleColor And &H80000000) Then
        OleToColor = GetSysColor(OleColor And &HFF&)
    Else
        OleToColor = OleColor
    End If
        
End Function


Private Sub PaintTransparentPicture(ByVal hDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
                                    

    ' Purpose:  Draws a transparent bitmap to a DC.  The pixels of the passed
    '           bitmap that match the passed mask color will not be painted
    '           to the destination DC
    ' In:
    '   [hdcDest]
    '           Device context to paint the picture on
    '   [xDest]
    '           X coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [yDest]
    '           Y coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [Width]
    '           Width of picture area to paint in pixels.  Note: If this value
    '           is outrageous (i.e.: you passed a forms ScaleWidth in twips
    '           instead of the pictures' width in pixels), this procedure will
    '           attempt to create bitmaps that require outrageous
    '           amounts of memory.
    '   [Height]
    '           Height of picture area to paint in pixels.  Note: If this
    '           value is outrageous (i.e.: you passed a forms ScaleHeight in
    '           twips instead of the pictures' height in pixels), this
    '           procedure will attempt to create bitmaps that require
    '           outrageous amounts of memory.
    '   [picSource]
    '           Standard Picture object to be used as the image source
    '   [xSrc]
    '           X coordinate of the upper left corner of the area in the picture
    '           to use as the source. (in pixels)
    '           Ignored if picSource is an Icon.
    '   [ySrc]
    '           Y coordinate of the upper left corner of the area in the picture
    '           to use as the source. (in pixels)
    '           Ignored if picSource is an Icon.
    '   [clrMask]
    '           Color of pixels to be masked out
    '   [hPal]
    '           Handle of palette to select into the memory DC's used to create
    '           the painting effect.
    '           If not provided, a HalfTone palette is used.


    Dim hdcSrc          As Long         'hDC that the source bitmap is selected into
    Dim hbmMemSrcOld    As Long
    Dim hbmMemSrc       As Long
    Dim udtRect         As RECT
    Dim hbrMask         As Long
    Dim lMaskColor      As Long
    Dim hDCScreen       As Long
    Dim hPalOld         As Long
    
    
    ' Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then
        
        Exit Sub
    End If

    If picSource.Type = vbPicTypeBitmap Then
        'Create halftone palette
        hDCScreen = GetDC(0&)
        m_hpalHalftone = CreateHalftonePalette(hDCScreen)
        ' Validate palette
        If hPal = 0 Then
            hPal = m_hpalHalftone
        End If
        hdcSrc = CreateCompatibleDC(hDCScreen)
        
        ' Select passed picture into an hDC
        hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
        hPalOld = SelectPalette(hdcSrc, hPal, True)
        RealizePalette hdcSrc
        ' Draw the bitmap
        PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
        SelectObject hdcSrc, hbmMemSrcOld
    
        ' Clean up
        SelectPalette hdcSrc, hPalOld, True
        RealizePalette hdcSrc
        DeleteDC hdcSrc
        ReleaseDC 0&, hDCScreen
        DeleteObject m_hpalHalftone
    End If
        
End Sub

Private Sub PaintTransparentDC(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
                                    
                                    
    ' Purpose:  Draws a transparent bitmap to a DC.  The pixels of the passed
    '           bitmap that match the passed mask color will not be painted
    '           to the destination DC
    '
    ' Called by:    PaintTransparentPicture()
    '
    ' In:
    '   [hdcDest]
    '           Device context to paint the picture on
    '   [xDest]
    '           X coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [yDest]
    '           Y coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [Width]
    '           Width of picture area to paint in pixels.  Note: If this value
    '           is outrageous (i.e.: you passed a forms ScaleWidth in twips
    '           instead of the pictures' width in pixels), this procedure will
    '           attempt to create bitmaps that require outrageous
    '           amounts of memory.
    '   [Height]
    '           Height of picture area to paint in pixels.  Note: If this
    '           value is outrageous (i.e.: you passed a forms ScaleHeight in
    '           twips instead of the pictures' height in pixels), this
    '           procedure will attempt to create bitmaps that require
    '           outrageous amounts of memory.
    '   [hdcSrc]
    '           Device context that contains the source picture
    '   [xSrc]
    '           X coordinate of the upper left corner of the area in the picture
    '           to use as the source. (in pixels)
    '   [ySrc]
    '           Y coordinate of the upper left corner of the area in the picture
    '           to use as the source. (in pixels)
    '   [clrMask]
    '           Color of pixels to be masked out
    '   [hPal]
    '           Handle of palette to select into the memory DC's used to create
    '           the painting effect.
    '           If not provided, a HalfTone palette is used.
                                    
                                    
                                    
    Dim hdcMask         As Long     ' hDC of the created mask image
    Dim hdcColor        As Long     ' hDC of the created color image
    Dim hbmMask         As Long     ' Bitmap handle to the mask image
    Dim hbmColor        As Long     ' Bitmap handle to the color image
    Dim hbmColorOld     As Long
    Dim hbmMaskOld      As Long
    Dim hPalOld         As Long
    Dim hDCScreen       As Long
    Dim hdcScnBuffer    As Long     ' Buffer to do all work on
    Dim hbmScnBuffer    As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld   As Long
    Dim lMaskColor      As Long
    
    
    hDCScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    
    ' Create a color bitmap to server as a copy of the destination
    ' Do all work on this bitmap and then copy it back over the destination when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hDCScreen, Width, Height)
    ' Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hDCScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    ' Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy
    
    ' Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    ' hdcSrc, because this will create a DIB section if the original bitmap is a DIB section)
    hbmColor = CreateCompatibleBitmap(hDCScreen, Width, Height)
    ' Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    ' First, blt the source bitmap onto the cover.  We do this first
    ' and then use it instead of the source bitmap
    ' because the source bitmap may be
    ' a DIB section, which behaves differently than a bitmap.
    ' (Specifically, copying from a DIB section to a monochrome bitmap
    ' does a nearest-color selection rather than painting based on the
    ' backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hDCScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    ' In case hdcSrc contains a monochrome bitmap, we must set the destination
    ' foreground/background colors according to those currently set in hdcSrc
    ' (because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hDCScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    ' When bitblt'ing from color to monochrome, Windows sets to 1
    ' all pixels that match the background color of the source DC.  All
    ' other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    ' Paint the rest of the cover bitmap.
    '
    ' What we want here is black at the transparent color, and
    ' the original colors everywhere else.  To do this, we first
    ' paint the original onto the cover (which we already did), then we
    ' AND the inverse of the mask onto that using the DSna ternary raster
    ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    ' Operation Codes", "Ternary Raster Operations", or search in MSDN
    ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    ' When bitblt'ing from monochrome to color, Windows transforms all white
    ' bits (1) to the background color of the destination hDC.  All black (0)
    ' bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    ' Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    ' Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    ' Copy the screen buffer to the screen
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    ' All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hDCScreen

End Sub



' *************************************
' *           PROPERTIES              *
' *************************************

Public Property Get BackColor() As OLE_COLOR

    BackColor = oBackColor
    
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
        
    Select Case m_Apperance
    
    Case [XP Diagonal Left], [XP Diagonal Right]
            oBackColor = UserControl.Parent.BackColor
    
    Case Else
            oBackColor = NewBackColor
            
    End Select
    
    UserControl.BackColor = oBackColor
    DrawStatusBar True
    PropertyChanged "BackColor"
 
End Property


Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = oForeColor
 
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    
    oForeColor = NewForeColor
    UserControl.ForeColor = oForeColor
    DrawStatusBar True
    PropertyChanged "ForeColor"
    
End Property


Public Property Get NumberOfPanels() As Long
    
    NumberOfPanels = m_PanelCount
 
End Property


Public Property Get PanelWidth(ByVal Index As Long) As Long
    
    PanelWidth = m_Panels(Index).ClientWidth
 
End Property

Public Property Let PanelWidth(ByVal Index As Long, ByVal NewPanelWidth As Long)
    
    If NewPanelWidth < 1 Then
    
        Exit Property
    End If
    m_Panels(Index).ClientWidth = NewPanelWidth
    DrawStatusBar True
    PropertyChanged "PWidth"
 
End Property


Public Property Get PanelMinWidth(ByVal Index As Long) As Long
    
    PanelMinWidth = m_Panels(Index).pMinWidth
 
End Property

Public Property Let PanelMinWidth(ByVal Index As Long, ByVal NewPanelMinWidth As Long)
    
    If NewPanelMinWidth < 1 Then
    
        Exit Property
    End If
    m_Panels(Index).pMinWidth = NewPanelMinWidth
    DrawStatusBar True
    PropertyChanged "PMinWidth"
 
End Property


Public Property Get PanelCaption(ByVal Index As Long) As String
    
    PanelCaption = m_Panels(Index).PanelText
 
End Property

Public Property Let PanelCaption(ByVal Index As Long, ByVal NewPanelCaption As String)
    
    Dim rc As RECT
    
    If Index < 1 Or Index > m_PanelCount Then
                        ' Maybe we should raise an error here ... ?
        Exit Property
    End If
    
    ClearPanel Index
    
    With m_Panels(Index)
        .PanelText = NewPanelCaption
    
        ' Draw the new contents
        rc.lLeft = .ContentsLeft
        rc.lTop = .ContentsTop
        rc.lRight = .ContentsRight
        rc.lBottom = .ContentsBottom
    
        SetTextColor UserControl.hDC, IIf(.pEnabled = True, oForeColor, oDissColor)
        OffsetRect rc, 4, (ScaleHeight - rc.lBottom) / 2
        DrawTheText UserControl.hDC, .PanelText, Len(.PanelText), rc, .TextAlignment
    End With
    
    UserControl.Refresh
    
    PropertyChanged "pText"
    
End Property


Public Property Get ToolTipText(ByVal Index As Long) As String
    
    ToolTipText = m_Panels(Index).ToolTipTxt
    
End Property

Public Property Let ToolTipText(ByVal Index As Long, ByVal NewToolTipText As String)
    
    m_Panels(Index).ToolTipTxt = NewToolTipText
    PropertyChanged "pTTText"
    
End Property


Public Property Get PanelPicture(ByVal Index As Long) As StdPicture
    
    Set PanelPicture = m_Panels(Index).PanelPicture
    
End Property

Public Property Set PanelPicture(ByVal Index As Long, ByVal NewPanelPicture As StdPicture)
    
    Set m_Panels(Index).PanelPicture = NewPanelPicture
    DrawStatusBar False
    PropertyChanged "PanelPicture"
    
End Property


Public Property Get PanelEnabled(ByVal Index As Long) As Boolean
    
    PanelEnabled = m_Panels(Index).pEnabled
    
End Property

Public Property Let PanelEnabled(ByVal Index As Long, ByVal NewEnabled As Boolean)
    
    m_Panels(Index).pEnabled = NewEnabled
    DrawStatusBar False
    PropertyChanged "pEnabled"
    
End Property


Public Property Get PanelVisible(ByVal Index As Long) As Boolean
    
    PanelVisible = m_Panels(Index).pVisible
    
End Property

Public Property Let PanelVisible(ByVal Index As Long, ByVal NewVisible As Boolean)
    
    m_Panels(Index).pVisible = NewVisible
    DrawStatusBar True
    PropertyChanged "pVisible"
    
End Property


Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Used for panel pics / icons. Set BEFORE loading an icon!"
    
    MaskColor = oMaskColor
    
End Property

Public Property Let MaskColor(ByVal NewMaskColor As OLE_COLOR)
    
    oMaskColor = NewMaskColor
    DrawStatusBar False
    PropertyChanged "MaskColor"
    
End Property


Public Property Get Font() As Font

    Set Font = UserControl.Font
    
End Property

Public Property Set Font(ByVal NewFont As Font)

    Set UserControl.Font = NewFont
    DrawStatusBar False
    PropertyChanged "Font"
    
End Property


Public Property Get ShowGripper() As Boolean
Attribute ShowGripper.VB_Description = "Draw form size gripper in bottom right corner of statusbar."

    ShowGripper = bDrawGripper
    
End Property

Public Property Let ShowGripper(ByVal newValue As Boolean)

    bDrawGripper = newValue
    DrawStatusBar True
    If bDrawGripper = True Then
        With UserControl
            If TypeOf .Parent Is Form Then
                If Not TypeOf .Parent Is MDIForm Then
                    Set frm = .Parent
                    If Ambient.UserMode Then
                        Set eForm = frm
                    End If
                End If
            End If
        End With
    Else
        ReleaseCapture
    End If
    PropertyChanged "ShowGripper"
    
End Property


Public Property Get ShowSeperators() As Boolean
Attribute ShowSeperators.VB_Description = "Draw seperating lines between panels."

    ShowSeperators = bDrawSeperators
    
End Property

Public Property Let ShowSeperators(ByVal newValue As Boolean)

    bDrawSeperators = newValue
    DrawStatusBar True
    PropertyChanged "ShowSeperators"
    
End Property


Public Property Get ForeColorDisabled() As OLE_COLOR

    ForeColorDisabled = oDissColor
    
End Property

Public Property Let ForeColorDisabled(ByVal NewDissColor As OLE_COLOR)

    oDissColor = NewDissColor
    DrawStatusBar False
    PropertyChanged "ForeColorDissabled"
    
End Property


Public Property Get PanelType(ByVal Index As Long) As enPanelType

    PanelType = m_Panels(Index).PanelType
    
End Property

Public Property Let PanelType(ByVal Index As Long, ByVal NewPanelType As enPanelType)

    m_Panels(Index).PanelType = NewPanelType
    DrawStatusBar False
    PropertyChanged "pType"
    
End Property


Public Property Get TextAlignment(ByVal Index As Long) As enTextAlignment

    TextAlignment = m_Panels(Index).TextAlignment
    
End Property

Public Property Let TextAlignment(ByVal Index As Long, ByVal NewTextAlignment As enTextAlignment)

    m_Panels(Index).TextAlignment = NewTextAlignment
    DrawStatusBar False
    PropertyChanged "pTextAlignment"
    
End Property


Public Property Get PanelPicAlignment(ByVal Index As Long) As enPanelPictureAlignment

    PanelPicAlignment = m_Panels(Index).PanelPicAlignment
    
End Property

Public Property Let PanelPicAlignment(ByVal Index As Long, ByVal NewPanelPicAlignment As enPanelPictureAlignment)

    m_Panels(Index).PanelPicAlignment = NewPanelPicAlignment
    DrawStatusBar False
    PropertyChanged "pPAlignment"
    
End Property


Public Property Get PanelBckgColor(ByVal Index As Long) As Long

    PanelBckgColor = m_Panels(Index).PanelBckgColor
    
End Property

Public Property Let PanelBckgColor(ByVal Index As Long, ByVal NewPanelBckgColor As Long)

    m_Panels(Index).PanelBckgColor = NewPanelBckgColor
    DrawStatusBar False
    PropertyChanged "pBckgColor"
    
End Property


Public Property Get PanelGradient(ByVal Index As Long) As Long

    PanelGradient = m_Panels(Index).PanelGradient
    
End Property

Public Property Let PanelGradient(ByVal Index As Long, ByVal NewPanelGradient As Long)

    m_Panels(Index).PanelGradient = NewPanelGradient
    DrawStatusBar False
    PropertyChanged "pGradient"
    
End Property


Public Property Get PanelEdgeSpacing(ByVal Index As Long) As Long

    PanelEdgeSpacing = m_Panels(Index).PanelEdgeSpacing
    
End Property

Public Property Let PanelEdgeSpacing(ByVal Index As Long, ByVal NewPanelEdgeSpacing As Long)

    m_Panels(Index).PanelEdgeSpacing = NewPanelEdgeSpacing
    DrawStatusBar False
    PropertyChanged "pEdgeSpacing"
    
End Property


Public Property Get PanelEdgeInner(ByVal Index As Long) As Long

    PanelEdgeInner = m_Panels(Index).PanelEdgeInner
    
End Property

Public Property Let PanelEdgeInner(ByVal Index As Long, ByVal NewPanelEdgeInner As Long)

    m_Panels(Index).PanelEdgeInner = NewPanelEdgeInner
    DrawStatusBar False
    PropertyChanged "pEdgeInner"
    
End Property


Public Property Get PanelEdgeOuter(ByVal Index As Long) As Long

    PanelEdgeOuter = m_Panels(Index).PanelEdgeOuter
    
End Property

Public Property Let PanelEdgeOuter(ByVal Index As Long, ByVal NewPanelEdgeOuter As Long)

    m_Panels(Index).PanelEdgeOuter = NewPanelEdgeOuter
    DrawStatusBar False
    PropertyChanged "pEdgeOuter"
    
End Property


Public Property Get Apperance() As enVWsbXPApperance
Attribute Apperance.VB_Description = "Select styling of statusbar."
   
    Apperance = m_Apperance

End Property

Public Property Let Apperance(ByVal New_Apperance As enVWsbXPApperance)
   
    m_Apperance = New_Apperance
    DrawStatusBar True
    PropertyChanged "Apperance"

    Select Case m_Apperance

        Case [XP Diagonal Left], [XP Diagonal Right]
            BackColor = UserControl.Ambient.BackColor

    End Select

End Property


Public Property Get UseWindowsColors() As Boolean
Attribute UseWindowsColors.VB_Description = "Try to use windows default colors. (Not fully implemented yet, sorry!)"
   
    UseWindowsColors = m_UseWindowsColors

End Property

Public Property Let UseWindowsColors(ByVal New_UseWindowsColors As Boolean)
   
    m_UseWindowsColors = New_UseWindowsColors
    DrawStatusBar True
    PropertyChanged "UseWindowsColors"

End Property


Property Get ShowTopLine() As Boolean
Attribute ShowTopLine.VB_Description = "Draw a line on top of the statusbar. Color depence on Apperance."
    
    ShowTopLine = m_TopLine
    
End Property

Property Let ShowTopLine(ByVal New_ShowTopLine As Boolean)
    
    m_TopLine = New_ShowTopLine
    DrawStatusBar True
    PropertyChanged "TopLine"
    
End Property


Public Property Get PanelTag(ByVal Index As Long) As Variant

    PanelTag = m_Panels(Index).Tag
    
End Property

Public Property Let PanelTag(ByVal Index As Long, ByVal NewPanelTag As Variant)

    m_Panels(Index).Tag = NewPanelTag
    PropertyChanged "pTag"
        
End Property


' #*#







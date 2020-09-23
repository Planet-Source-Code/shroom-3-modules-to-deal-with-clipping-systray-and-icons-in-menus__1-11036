Attribute VB_Name = "RestrictSizing"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function ClipCursor Lib "user32.dll" (lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClipCursor Lib "user32.dll" (lprc As RECT) As Long
Dim frmHeight As Long, frmWidth As Long

Public Function GetX()
    Dim Point As POINTAPI, RetVal As Long
    RetVal = GetCursorPos(Point)
    GetX = Point.X
End Function

Public Function GetY()
    Dim Point As POINTAPI, RetVal As Long
    RetVal = GetCursorPos(Point)
    GetY = Point.Y
End Function

Public Sub SetClipVars(MinHeight As Long, MinWidth As Long)
    frmHeight = MinHeight
    frmWidth = MinWidth
End Sub

Public Sub ClipForForm(frm As Form, MinHeight As Long, MinWidth As Long)
    Dim ResizeREC As RECT
    Dim DesktopREC As RECT
    Dim GetClipREC As RECT
    ResizeREC.Top = (frm.Top + MinHeight) / Screen.TwipsPerPixelY - 2
    ResizeREC.Bottom = (frm.Top + (frm.Height - MinHeight)) / Screen.TwipsPerPixelY + 2
    ResizeREC.Left = (frm.Left + MinWidth) / Screen.TwipsPerPixelX - 2
    ResizeREC.Right = (frm.Left + (frm.Width - MinWidth)) / Screen.TwipsPerPixelX + 2
    If frm.Width <> frmWidth And frm.WindowState = 0 Then
        RetVal = GetClipCursor(GetClipREC)
        DeskhWnd = GetDesktopWindow()
        RetVal = GetWindowRect(DeskhWnd, DesktopREC)
        If GetX > ((frm.Left + (MinWidth / 2)) / Screen.TwipsPerPixelX) And GetClipREC.Right = DesktopREC.Right Then
            If GetClipREC.Left = ResizeREC.Left And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Left = ResizeREC.Left And (GetClipREC.Top <> ResizeREC.Top Or GetClipREC.Bottom <> ResizeREC.Bottom) Then
                If frm.Height <> frmHeight Then
                    If GetY > (frm.Top / Screen.TwipsPerPixelY) + 25 Then
                        DesktopREC.Top = ResizeREC.Top
                    Else
                        DesktopREC.Bottom = ResizeREC.Bottom
                    End If
                    frmHeight = frm.Height
                End If
            ElseIf GetClipREC.Left <> ResizeREC.Left And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then
                If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
                If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
                DesktopREC.Left = ResizeREC.Left
            End If
            DesktopREC.Left = ResizeREC.Left
        Else
            If GetClipREC.Right = ResizeREC.Right And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Right = ResizeREC.Right And (GetClipREC.Bottom <> ResizeREC.Bottom Or GetClipREC.Top <> ResizeREC.Top) Then
                If frm.Height <> frmHeight Then
                    If GetY > (frm.Top / Screen.TwipsPerPixelY) + 25 Then
                        DesktopREC.Top = ResizeREC.Top
                    Else
                        DesktopREC.Bottom = ResizeREC.Bottom
                    End If
                    frmHeight = frm.Height
                End If
            ElseIf GetClipREC.Right <> ResizeREC.Right And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then
                If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
                If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
            End If
            DesktopREC.Right = ResizeREC.Right
        End If
        RetVal = ClipCursor(DesktopREC)
        frmWidth = frm.Width
    ElseIf frm.Height <> frmHeight And frm.WindowState = 0 Then
        RetVal = GetClipCursor(GetClipREC)
        DeskhWnd = GetDesktopWindow()
        RetVal = GetWindowRect(DeskhWnd, DesktopREC)
        If GetY > ((frm.Top + (MinHeight / 2)) / Screen.TwipsPerPixelY) And GetClipREC.Bottom <> ResizeREC.Bottom Then
            If GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left = ResizeREC.Left Or GetClipREC.Right = ResizeREC.Right) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left <> ResizeREC.Left Or GetClipREC.Right <> ResizeREC.Right) Then
                If frm.Width <> frmWidth Then
                    If GetX > (frm.Left / Screen.TwipsPerPixelX) + 15 Then
                        DesktopREC.Left = ResizeREC.Left
                    Else
                        DesktopREC.Right = ResizeREC.Right
                    End If
                    frmWidth = frm.Width
                End If
            ElseIf GetClipREC.Top <> ResizeREC.Top And (GetClipREC.Left = ResizeREC.Left Or GetClipREC.Right = ResizeREC.Right) Then
                If GetClipREC.Left = ResizeREC.Left Then DesktopREC.Left = GetClipREC.Left
                If GetClipREC.Right = ResizeREC.Right Then DesktopREC.Right = GetClipREC.Right
            End If
            DesktopREC.Top = ResizeREC.Top
        Else
            If GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right = ResizeREC.Right Or GetClipREC.Left = ResizeREC.Left) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right <> ResizeREC.Right Or GetClipREC.Left <> ResizeREC.Left) Then
                If frm.Width <> frmWidth Then
                    If GetX > (frm.Left / Screen.TwipsPerPixelX) + 15 Then
                        DesktopREC.Left = ResizeREC.Left
                    Else
                        DesktopREC.Right = ResizeREC.Right
                    End If
                    frmWidth = frm.Width
                End If
            ElseIf GetClipREC.Bottom <> ResizeREC.Bottom And (GetClipREC.Right = ResizeREC.Right Or GetClipREC.Left = ResizeREC.Left) Then
                If GetClipREC.Left = ResizeREC.Left Then DesktopREC.Left = GetClipREC.Left
                If GetClipREC.Right = ResizeREC.Right Then DesktopREC.Right = GetClipREC.Right
            End If
            DesktopREC.Bottom = ResizeREC.Bottom
        End If
        RetVal = ClipCursor(DesktopREC)
        frmHeight = frm.Height
    End If
End Sub

Public Sub RemoveClipping()
    Dim DesktopREC As RECT, RetVal As Long, DeskhWnd As Long
    DeskhWnd = GetDesktopWindow()
    RetVal = GetWindowRect(DeskhWnd, DesktopREC)
    RetVal = ClipCursor(DesktopREC)
End Sub

Attribute VB_Name = "Module1"
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Public Enum MouseButtons
   LeftMouseButton
   RightMouseButton
   MiddleMouseButton
End Enum

Public Sub MouseUp(MouseButton As MouseButtons)
   Select Case (MouseButton)
      Case LeftMouseButton
         Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
      Case MiddleMouseButton
         Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
      Case RightMouseButton
         Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
   End Select
End Sub

Public Sub MouseDown(MouseButton As MouseButtons)
   Select Case (MouseButton)
      Case LeftMouseButton
         Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
      Case MiddleMouseButton
         Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
      Case RightMouseButton
         Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
   End Select
End Sub

Public Sub MouseClick(MouseButton As MouseButtons)
   MouseDown (MouseButton)
   MouseUp (MouseButton)
End Sub


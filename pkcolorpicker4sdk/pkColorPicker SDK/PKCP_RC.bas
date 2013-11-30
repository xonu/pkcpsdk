Attribute VB_Name = "PKCP"
' Remote Control 1.0 for pkColorPicker 4
' Copyright (c) 2003 Pawel Kazakow
' Homepage: http://www.pkworld.de
' Email: info@pkworld.de

'------- ------ ----- ---- --- -- -
Option Explicit

Private Const DebugingMode = 0
'------- ------ ----- ---- --- -- -

' pkColorPicker registry path
Private Const PKCP_RegPath = "Software\PKSOFT\pkColorPicker\4.00"

' pkColorPicker messages
Public PKCP_STATE As Long
Public PKCP_SETCOLOR As Long
Public PKCP_GETCOLOR As Long

' pkColorPicker state constants
Public Const PKCP_QUIT = 0&
Public Const PKCP_SHOW = 1&
Public Const PKCP_HIDE = 2&

'------- ------ ----- ---- --- -- -

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001

'------- ------ ----- ---- --- -- -

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'------- ------ ----- ---- --- -- -

Public Sub Init()
 RegisterMessages
End Sub

Public Function hWnd() As Long
 hWnd = GetSettingLong(HKEY_CURRENT_USER, PKCP_RegPath, "hWnd", 0)
End Function

Public Function Launch(Optional Arguments As String = "") As Long
 On Error Resume Next
 Launch = CBool(Shell(Chr$(34) & GetSettingString(HKEY_CURRENT_USER, PKCP_RegPath, "Executable", "") & Chr$(34) & " " & Arguments, vbNormalFocus))
End Function

'------- ------ ----- ---- --- -- -

Public Property Get State() As Long
 If IsWindow(hWnd()) Then
  State = SendMessage(hWnd(), PKCP_STATE, 0&, 0&)
 Else
  State = PKCP_QUIT
 End If
End Property

Public Property Let State(ByVal vNewValue As Long)
 Call SendMessage(hWnd(), PKCP_STATE, vNewValue, -1&)
End Property


Public Function Quit() As Boolean
 Quit = SendMessage(hWnd(), PKCP_STATE, PKCP_QUIT, -1&)
End Function

Public Function Show() As Boolean
 Show = SendMessage(hWnd(), PKCP_STATE, PKCP_SHOW, -1&)
End Function

Public Function Hide() As Boolean
 Hide = SendMessage(hWnd(), PKCP_STATE, PKCP_HIDE, -1&)
End Function

'------- ------ ----- ---- --- -- -

Public Property Get Color1() As Long
 Color1 = SendMessage(hWnd(), PKCP_GETCOLOR, -1&, 0&)
End Property

Public Property Let Color1(ByVal NewColor As Long)
 Call SendMessage(hWnd(), PKCP_SETCOLOR, NewColor, Color2)
End Property


Public Property Get Color2() As Long
 Color2 = SendMessage(hWnd(), PKCP_GETCOLOR, 0&, -1&)
End Property

Public Property Let Color2(ByVal NewColor As Long)
 Call SendMessage(hWnd(), PKCP_SETCOLOR, Color1, NewColor)
End Property


Public Function SetColors(Color1 As Long, Color2 As Long) As Long
 SetColors = SendMessage(hWnd(), PKCP_SETCOLOR, Color1, Color2)
End Function

'------- ------ ----- ---- --- -- -
' private functions

Private Sub RegisterMessages()
 PKCP_STATE = RegisterWindowMessage("PKCP_STATE")
 PKCP_SETCOLOR = RegisterWindowMessage("PKCP_SETCOLOR")
 PKCP_GETCOLOR = RegisterWindowMessage("PKCP_GETCOLOR")
End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
   Dim hCurKey As Long
   Dim lValueType As Long
   Dim strBuffer As String
   Dim lDataBufferSize As Long
   Dim intZeroPos As Integer
   Dim lRegResult As Long
   
   If Not IsEmpty(Default) Then
     GetSettingString = Default
   Else
     GetSettingString = ""
   End If
   lRegResult = RegOpenKey(hKey, strPath, hCurKey)
   lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
   If lRegResult = 0 Then
      If lValueType = 1 Then
         strBuffer = String(lDataBufferSize, " ")
         lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
         intZeroPos = InStr(strBuffer, Chr$(0))
         If intZeroPos > 0 Then
            GetSettingString = Left$(strBuffer, intZeroPos - 1)
         Else
            GetSettingString = strBuffer
         End If
      End If
   End If
   lRegResult = RegCloseKey(hCurKey)
End Function

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
   Dim lRegResult As Long
   Dim lValueType As Long
   Dim lBuffer As Long
   Dim lDataBufferSize As Long
   Dim hCurKey As Long
   
   If Not IsEmpty(Default) Then
     GetSettingLong = Default
   Else
     GetSettingLong = 0
   End If
   lRegResult = RegOpenKey(hKey, strPath, hCurKey)
   lDataBufferSize = 4
   lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
   If lRegResult = 0 Then
      If lValueType = 4 Then
         GetSettingLong = lBuffer
      End If
   End If
   lRegResult = RegCloseKey(hCurKey)
End Function

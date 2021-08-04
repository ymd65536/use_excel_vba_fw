Attribute VB_Name = "UseWin32API"
Option Explicit

' �ϐ��A�萔�A�֐����̖����Ƀf�[�^�^�������L����t���邱�Ƃ�
' �^���w�肷�邱�Ƃ��\
'
' % ���� integer
' & ������ Long
' ! �P���x���������_�^ Single
' # �{���x���������_�^ Double
' $ ������^ String
'

Private Const WM_SYSCHAR = &H106
Public Const WM_COMMAND As Long = &H111&
Public Const SC_CLOSE = &HF060
Public Const WM_SYSCOMMAND = &H112
Public Const WM_CLOSE = &H10

Public Const SW_SHOWMINIMIZED = 2   '�A�N�e�B�u�ɂ��čŏ���
Public Const SW_MINIMIZE = 6        '�ŏ���
Public Const SW_HIDE = 0            '��\��


'SetWindowPos��hWndInsertAfter�Ŏw�肷��l�̒�`
Public Const HWND_TOP = 0           '��O�ɃZ�b�g
Public Const HWND_BOTTOM = 1        '���ɃZ�b�g
Public Const HWND_TOPMOST = -1    '��Ɏ�O�ɃZ�b�g
Public Const HWND_NOTOPMOST = -2    '��Ɏ�O������

'SetWindowPos��wFlags�Ɏw�肷��l�̒�`
Const SWP_SHOWWINDOW = &H40  '�E�B���h�E��\������
Public Const SWP_NOSIZE = &H1   '�E�B���h�E�̃T�C�Y��ύX���Ȃ�
Public Const SWP_NOMOVE = &H2   '�E�B���h�E�̈ʒu��ύX���Ȃ�

'
' �ҋ@���ԃ~���b�Ŏw�肷��B
'
Public Declare Sub Sleep Lib _
"kernel32" (ByVal dwMilliseconds As Long)

' user32.dll
' �E�B���h�E���ŏ�������API
'
Public Declare Function ShowWindow _
  Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) _
  As Long
  
' user32.dll
' �E�B���h�E��T��API
'
Public Declare Function FindWindow _
  Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) _
  As LongPtr


Public Declare Function FindWindowEx _
Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr


' user32.dll
' �E�B���h�E�����݂��邩���m�F����
'
Public Declare Function IsWindowVisible _
  Lib "user32" (ByVal hWnd As LongPtr) As Long

' user32.dll
' �A�v���P�[�V�����Ƀ��b�Z�[�W�𑗐M(POST)
'

Public Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" _
(ByVal hWnd As LongPtr, ByVal wMsg As Long, _
ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

'�E�B���h�E�̃T�C�Y�A�ʒu�A�����Z�I�[�_�[��ݒ�i�w�ʂƂ��O�ʂƂ��j
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' user32.dll
' �A�v���P�[�V�����Ƀ��b�Z�[�W�𑗐M
'
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hWnd&, _
ByVal wMsg&, _
ByVal wParam&, _
ByVal lParam&) As Long

'
' ��ʂ̈ʒu��ύX
'
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long


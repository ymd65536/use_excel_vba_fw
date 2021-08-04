Attribute VB_Name = "UseWin32API"
Option Explicit

' 変数、定数、関数名の末尾にデータ型を示す記号を付けることで
' 型を指定することが可能
'
' % 整数 integer
' & 長整数 Long
' ! 単精度浮動小数点型 Single
' # 倍精度浮動小数点型 Double
' $ 文字列型 String
'

Private Const WM_SYSCHAR = &H106
Public Const WM_COMMAND As Long = &H111&
Public Const SC_CLOSE = &HF060
Public Const WM_SYSCOMMAND = &H112
Public Const WM_CLOSE = &H10

Public Const SW_SHOWMINIMIZED = 2   'アクティブにして最小化
Public Const SW_MINIMIZE = 6        '最小化
Public Const SW_HIDE = 0            '非表示


'SetWindowPosのhWndInsertAfterで指定する値の定義
Public Const HWND_TOP = 0           '手前にセット
Public Const HWND_BOTTOM = 1        '後ろにセット
Public Const HWND_TOPMOST = -1    '常に手前にセット
Public Const HWND_NOTOPMOST = -2    '常に手前を解除

'SetWindowPosのwFlagsに指定する値の定義
Const SWP_SHOWWINDOW = &H40  'ウィンドウを表示する
Public Const SWP_NOSIZE = &H1   'ウィンドウのサイズを変更しない
Public Const SWP_NOMOVE = &H2   'ウィンドウの位置を変更しない

'
' 待機時間ミリ秒で指定する。
'
Public Declare Sub Sleep Lib _
"kernel32" (ByVal dwMilliseconds As Long)

' user32.dll
' ウィンドウを最小化するAPI
'
Public Declare Function ShowWindow _
  Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) _
  As Long
  
' user32.dll
' ウィンドウを探すAPI
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
' ウィンドウが存在するかを確認する
'
Public Declare Function IsWindowVisible _
  Lib "user32" (ByVal hWnd As LongPtr) As Long

' user32.dll
' アプリケーションにメッセージを送信(POST)
'

Public Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" _
(ByVal hWnd As LongPtr, ByVal wMsg As Long, _
ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

'ウィンドウのサイズ、位置、およびZオーダーを設定（背面とか前面とか）
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' user32.dll
' アプリケーションにメッセージを送信
'
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hWnd&, _
ByVal wMsg&, _
ByVal wParam&, _
ByVal lParam&) As Long

'
' 画面の位置を変更
'
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long


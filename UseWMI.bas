Attribute VB_Name = "UseWMI"
Option Explicit

' Windows Management Instrumentation (WMI) �𗘗p
' https://www.atmarkit.co.jp/icd/root/21/81148321.html
' Windows�̃V�X�e���Ǘ��p�C���^�[�t�F�C�X�B
' �G���^�[�v���C�Y�E���x���̃V�X�e���Ǘ���ړI�Ƃ��čl�Ă��ꂽ�W���d�l��WBEM�i�ƊE�c�̂�DMTF������j�ɏ]����
' �}�C�N���\�t�g�������Windows OS��Ɏ����A�g���������́B
' WMI���񋟂���C���^�[�t�F�C�X�𗘗p����V�X�e���Ǘ��p�R���|�[�l���g�ɃA�N�Z�X���邱�Ƃ��ł���

' MAC�A�h���X���擾

Public Sub getMacAddress()
  Dim objNetwork As Object 'Windows�̏��
  Dim strNetworkSql As String 'Windows�̏��擾�� �ۑ��ϐ�
  Dim strMacAdr As String '�擾����MAC�A�h���X����

  'Windows�̏��擾�� �g�ݗ���
  strNetworkSql = "SELECT * FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL"

  'Windows�̏��擾�����g�������擾(1�ڂ̂�)
  For Each objNetwork In GetObject("winmgmts:").ExecQuery(strNetworkSql)
    strMacAdr = objNetwork.MACAddress
  Next
  
  '���b�Z�[�W�{�b�N�X��MAC�A�h���X��\��
  MsgBox (strMacAdr)

End Sub

Sub Sample()
  
  ''���ׂĂ̎������s���擾����
  Dim objWMIService, colStartupCommands, objStartupCommand
  Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
  Set colStartupCommands = objWMIService.ExecQuery("Select * from Win32_StartupCommand")
  
  For Each objStartupCommand In colStartupCommands
    Debug.Print "Command: " & objStartupCommand.Command
    Debug.Print "Description: " & objStartupCommand.Description
    Debug.Print "Location: " & objStartupCommand.Location
    Debug.Print "Name: " & objStartupCommand.Name
    Debug.Print "User: " & objStartupCommand.User
    Debug.Print ""
  Next
  Set colStartupCommands = Nothing
  Set objWMIService = Nothing
End Sub


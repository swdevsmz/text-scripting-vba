VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ����������
Private Sub Class_Initialize()
    Name = "Bob"
    Age = 30
    Set Mother = Nothing

    Debug.Print "initalized"
End Sub

' �I��������
Private Sub Class_Terminate()
    Name = ""
    Age = 0
    Set Mother = Nothing

    ' ���̑��A�g�p���Ă��郊�\�[�X�̉���������s��
    '

    Debug.Print "terminated"
End Sub
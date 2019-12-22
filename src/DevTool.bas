Attribute VB_Name = "DevTool"
Option Explicit

'<summary>
' �@�\: �SVBA�\�[�X���G�N�X�|�[�g����
' ����: �Ȃ�
' </summary>
' <remarks>
' 1. VBE�ɂ�����Microsoft Visual Basic for Applications Extensibility�ւ̎Q�Ƃ�ǉ�����B
' 2. �uVBA�v���W�F�N�g �I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�I�v�V�������w�肷��B
' </remarks>
Public Sub ExportAllSource()
    Dim module                  As VBComponent      ' ���W���[��
    Dim moduleList              As VBComponents     ' VBA�v���W�F�N�g�̑S���W���[��
    Dim extension               As String           ' ���W���[���̊g���q
    Dim sFilePath               As String           ' �G�N�X�|�[�g�t�@�C���p�X
    Dim sSaveFolder             As String           ' �ۑ���t�H���_
    
    sSaveFolder = ActiveWorkbook.Path & "\" & "src"

    If Dir(sSaveFolder, vbDirectory) = "" Then
        MkDir sSaveFolder
    End If
    
    ' �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = ActiveWorkbook.VBProject.VBComponents
    
    ' VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        ' �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ' �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
            ' .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        ' �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        ' ���̑�
        Else
            ' �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
        
        ' �G�N�X�|�[�g���{
        sFilePath = sSaveFolder & "\" & module.Name & "." & extension
        
        Call module.Export(sFilePath)
        
        ' �o�͐�m�F�p���O�o��
        Debug.Print sFilePath
CONTINUE:
    Next
    
End Sub
Attribute VB_Name = "Module_suzuki"
'---------------------------------------------------------------
'���ʏo�̓t�@�C���̏���������}�N��
'�i�P�j�G�N�Z���u�b�N���u�V�K�쐬�v����
'�i�Q�j�u�V�K�쐬�v�����G�N�Z���u�b�N���J���A�V�[�g���A�N�e�B�u�ɂ���
'�i�R�j���O��t���ĕۑ�����i�㏑���ۑ����m�F����j
'Owner suzuki
'---------------------------------------------------------------

Function OpenResultSheet(resultFile As Workbook, resultSheet As Worksheet)
    '�V�K�t�@�C���쐬
    Set resultFile = Workbooks.Add
    '�V�K�t�@�C����1 �Ԗڂ̃V�[�g��ϐ��Ɋi�[
    Set resultSheet = resultFile.Sheets(1)
 End Function
  
Function OutputResultFile(resultFile As Workbook)
    Dim resultFileName As String
 
   '�ۑ�����t�@�C���̖��O����́i���O��t���ĕۑ��j
      resultFileName = Application.GetSaveAsFilename( _
          InitialFileName:="�W�v����.xlsx", FileFilter:="Excel�t�@�C��, *.xlsx")

  '�t�@�C���������͂��ꂽ������
   If resultFileName <> "False" Then
  
     '�t�@�C���������͂��ꂽ�ꍇ
      '���O��t���ĕۑ�
       On Error GoTo Error1
       ActiveWorkbook.SaveAs Filename:=resultFileName
       Exit Function
    End If

'�G���[����
Error1:
    ActiveWorkbook.Close
    'ActiveWorkbook.Close savechanges:=False
    MsgBox "�ۑ�����܂���ł���"
    Err.Clear

End Function



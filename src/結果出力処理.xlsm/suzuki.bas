Attribute VB_Name = "suzuki"
'---------------------------------------------------------------
'�C���ʏo�̓t�@�C���̏���������}�N��
'�i�P�j�G�N�Z���u�b�N���u�V�K�쐬�v����
'�i�Q�j�u�V�K�쐬�v�����G�N�Z���u�b�N���J���A�V�[�g���A�N�e�B�u�ɂ���
'�i�R�j���O��t���ĕۑ�����i�㏑���ۑ����m�F����j
'Owner suzuki
'---------------------------------------------------------------
  
Sub OutputResultFile()
    Dim resultFileName As String
    Dim resultFile As Workbook
    Dim lRet As Long
    
   '�V�K�t�@�C���쐬
      Set resultFile = Workbooks.Add
   '�V�K�t�@�C����1 �Ԗڂ̃V�[�g���A�N�e�B�u
      resultFile.Sheets(1).Activate          '
 
   '�ۑ�����t�@�C���̖��O����́i���O��t���ĕۑ��j
      resultFileName = Application.GetSaveAsFilename( _
          InitialFilename:="�W�v����.xlsx", FileFilter:="Excel�t�@�C��, *.xlsx")

  '�t�@�C���������͂��ꂽ������
   If resultFileName <> "False" Then
  
     '�t�@�C���������͂��ꂽ�ꍇ
      '���O��t���ĕۑ�
       On Error GoTo Error1
       ActiveWorkbook.SaveAs Filename:=resultFileName
       Exit Sub
    End If

'�G���[����
Error1:
    ActiveWorkbook.Close
    'ActiveWorkbook.Close savechanges:=False
    MsgBox "�ۑ�����܂���ł���"
    Err.Clear

End Sub

Attribute VB_Name = "Module1"
'---------------------------------------------------------------
'�B���v���o�͂���}�N��
'�i�P�j�G�N�Z���V�[�g�w�肵�ĊJ��
'�i�Q�j����Q�̃Z���ɓ��͂��ꂽ�l���Q�̕ϐ��ɓǂݍ���
'�i�R�j���v�l���v�Z����
'�i�S�j���v�l��ʂ̃Z���ɏo�͂���
'Owner ooba
'---------------------------------------------------------------


 Sub Sumcells()

'-----�ϐ��錾-----
    Dim Path As String '�ΏۃG�N�Z���V�[�g�̃t�@�C���p�X
    
    Dim fcell As Double '�P�ڂ̒l
    Dim scell As Double '�Q�ڂ̒l
    
    Path = "C:\Users\xxxx\Desktop\hokan2\simnor2.xlsm" '���̃t�@�C���o�X
    
    '�i�P�j�G�N�Z���V�[�g�w�肵�ĊJ��
    Workbooks.Open Path
    
    '�i�Q�j����Q�̃Z���ɓ��͂��ꂽ�l���Q�̕ϐ��ɓǂݍ���
    fcell = Cells(1, 1) 'A1�̒l
    scell = Cells(1, 2) 'B1�̒l
    
    '�i�R�j���v�l���v�Z����
    '�i�S�j���v�l��ʂ̃Z���ɏo�͂���
    Cells(1, 3).Value = fcell + scell  'C1�ɍ��v�l���o��

    MsgBox Cells(1, 3) '�v�Z���ʏo��

 End Sub

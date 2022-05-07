// ****************************
// ��������
// ****************************
var App = new ActiveXObject( "Excel.Application" );
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel ��\��( ����������R�����g�� )
App.Visible = true;
// �x�����o���Ȃ�
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename ��O�ʂɏo����
// �{���A-4140 �ł��� WScript.Shell �� Run �Ɠ��� 2 ���g����
App.WindowState = 2

// �u�b�N�ǉ�
App.Workbooks.Add();

 // �u�b�N���擾( ������Ȃ��̂ŁACount �� 1 )
var Book = App.Workbooks( App.Workbooks.Count  );

Book.Sheets(1).Name = "�ŏ��̃V�[�g";
Book.Sheets.Add(null, Book.Sheets(1));
Book.Sheets(2).Name = "�ǉ��̃V�[�g";

// �擪�V�[�g���A�N�e�B�u�ɂ���
Book.Sheets(1).Activate();

// �Z���ɒl���Z�b�g
Book.Sheets(1).Cells(1, 1).Value = "�Ј���";
Book.Sheets(1).Cells(2, 1).Value = "�R�c�@���Y�r���q��";
Book.Sheets(1).Cells(3, 1).Value = "��؁@��Y";
Book.Sheets(1).Cells(4, 1).Value = "�����@�m�q";

// �񕝎�������
Book.Sheets(1).Columns("A:A").EntireColumn.AutoFit();

// ****************************
// �Q��
// �Ō�� 1 �́A�g�p����t�B���^�[
// �̔ԍ�
// ****************************
var FilePath = App.GetSaveAsFilename(null,"Excel �t�@�C�� (*.xlsx), *.xlsx", 1);
if ( FilePath == false ) {
    WshShell.Popup( "Excel �t�@�C���̕ۑ��I�����L�����Z������܂���" );
    App.Quit();
    App = null;
    WScript.Quit();
}

// ****************************
// �ۑ�
// �g���q�� .xls �ŕۑ�����ɂ�
// Book.SaveAs( FilePath, 56 )
// ****************************
try {
    Book.SaveAs( FilePath )
}
catch (error) {
    WshShell.Popup( "ERROR : " + error.description );
}

// �I��
App.Quit();

App = null;

WshShell.Run( "RunDLL32.EXE shell32.dll,ShellExec_RunDLL " + FilePath );

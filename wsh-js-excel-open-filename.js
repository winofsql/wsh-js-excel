var App = new ActiveXObject("Excel.Application");
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel ��\��( ����������R�����g�� )
App.Visible = true;
// �x�����o���Ȃ�
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename ��O�ʂɏo����
// �{���A-4140 �ł��� WScript.Shell �� Run �Ɠ��� 2 ���g����
App.WindowState = 2

// ��̃t�@�C�����J��
// https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.getopenfilename
var filePath = App.GetOpenFilename("�S��,*.*,CSV,*.csv", 1,"�t�@�C���̑I��",null, false );
// ���I���̏ꍇ
if( filePath === false ) {
    WshShell.Popup("�t�@�C���̎Q�ƑI�����L�����Z������܂���")
}
// �I���̏ꍇ
else {
    WshShell.Popup(filePath + " ��I�����܂���");
}

// �I��
App.Quit();

App = null;

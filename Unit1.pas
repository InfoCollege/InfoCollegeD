unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, DB, ADODB;

type
  TMainForm = class(TForm)
    Background: TImage;
    ProgramName: TLabel;
    L_University: TLabel;
    L_Auth: TLabel;
    Login: TEdit;
    Password: TEdit;
    L_Login: TLabel;
    L_Pass: TLabel;
    Auth: TButton;
    DB: TADOConnection;
    Query_Auth: TADOQuery;
    L_Version: TLabel;
    AuthDS1: TDataSource;
    procedure AuthClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses Unit2;

{$R *.dfm}

procedure TMainForm.AuthClick(Sender: TObject);
begin
Query_Auth.Close;
Query_Auth.SQL.Clear;
Query_Auth.SQL.Add('SELECT *');
Query_Auth.SQL.Add('FROM ���������� INNER JOIN ��������� ON ����������.��_���������=���������.��_���������');
Query_Auth.SQL.Add('WHERE �����=:P1');
Query_Auth.SQL.Add('AND ������=:P2;');
Query_Auth.Parameters.ParamByName('P1').Value:=Login.Text;
Query_Auth.Parameters.ParamByName('P2').Value:=Password.Text;
Query_Auth.Open;
//showmessage(ADOQuery1.SQL.text);
if Query_Auth.RecordCount = 1 then
begin
MainForm.Hide;
MenuChoice.show;
MenuChoice.username.Caption:=AuthDS1.DataSet.FindField('����_���').AsString;
end
else
showmessage('������ �1.����������� ������������ ���� �����/������');
end;

end.

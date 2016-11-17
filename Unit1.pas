unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, DB, ADODB;

type
  TForm1 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    Button1: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    Label7: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Unit2;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.Clear;
ADOQuery1.SQL.Add('SELECT *');
ADOQuery1.SQL.Add('FROM Сотрудники');
ADOQuery1.SQL.Add('WHERE Логин='''+Edit1.Text+'''');
ADOQuery1.SQL.Add('AND Пароль='''+Edit2.Text+''';');
ADOQuery1.Open;
//showmessage(ADOQuery1.SQL.text);
if ADOQuery1.RecordCount = 1 then
begin
Form1.Hide;
Form2.show;
Form2.Label2.Caption:=Form1.Edit1.Text;
end
else
showmessage('Ошибка №1.Обнаруженая неправильная пара логин/пароль');
end;

end.

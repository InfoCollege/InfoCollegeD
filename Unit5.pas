unit Unit5;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB;

type
  TForm5 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    ListBox1: TListBox;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit15: TEdit;
    Edit16: TEdit;
    Button2: TButton;
    Button1: TButton;
    ADOQuery1: TADOQuery;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form5: TForm5;
  sr:integer;
  sb:real;
  pred:integer;
implementation

{$R *.dfm}


procedure TForm5.Button1Click(Sender: TObject);
begin
pred:=strtoint(Edit10.Text)+strtoint(Edit11.Text)+strtoint(Edit12.Text);
 sr:=(5*strtoint(Edit10.Text))+(4*strtoint(Edit11.Text))+(3*strtoint(Edit12.Text));
 sb:=sr/pred;
 Edit13.Text:=FloatToStr(sb);
end;

procedure TForm5.Button2Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.Clear;
ADOQuery1.SQL.Add('INSERT INTO ПК(Фамилия,Имя,Отчество,НаименованиеОУ,Окончил,Паспорт,Датавыдачи,Выдан,Код,[Средний балл],Специальность,Телефон,[Телефон(2)],Примечание)');
ADOQuery1.SQL.Add('VALUES('''+Edit1.Text+''','''+Edit2.Text+''','''+Edit3.Text+''','''+Edit4.Text+''','''+Edit5.Text+''','''+Edit6.Text+''','''+Edit7.Text+''','''+Edit8.Text+''','''+Edit9.Text+''','''+Edit13.Text+'''');
ADOQuery1.SQL.Add(','''+Listbox1.Items[ListBox1.ItemIndex]+''','''+Edit14.Text+''','''+Edit15.Text+''','''+Edit16.Text+''');');
//showmessage(ADOQuery1.SQL.Text);
ADOQuery1.ExecSQL;
showmessage('Абитуриент внесен в реестр');
end;

end.

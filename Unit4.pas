unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, DBCtrls, DB, ADODB, Grids, DBGrids;

type
  TForm4 = class(TForm)
    Button2: TButton;
    Image1: TImage;
    Label1: TLabel;
    SD: TEdit;
    Label2: TLabel;
    ND: TEdit;
    Label3: TLabel;
    DV: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Fam: TEdit;
    Imya: TEdit;
    Otch: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Kval: TEdit;
    Prof: TListBox;
    Label9: TLabel;
    Label10: TLabel;
    DBGrid1: TDBGrid;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    Label11: TLabel;
    procedure ProfClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

uses Unit1,unit2;

{$R *.dfm}



procedure TForm4.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form4.Hide;
Form2.Show;
end;

procedure TForm4.ProfClick(Sender: TObject);
begin
if Prof.Selected[0]= true then
Kval.Text:='Техник по компьютерным системам';
if Prof.Selected[1]= true then
Kval.Text:='Техник-программист';
if Prof.Selected[2]= true then
Kval.Text:='Техник-программист';
if Prof.Selected[3]= true then
Kval.Text:='Техник по защите информации';
if Prof.Selected[4]= true then
Kval.Text:='Специалист по земельно-имущественным отношениям';
if Prof.Selected[5]= true then
Kval.Text:='Специалист по рекламе';
end;

procedure TForm4.Button2Click(Sender: TObject);
begin
if Form2.Label2.Caption='director'
then
  if  (SD.Text ='') or (ND.Text='') or (DV.Text ='') or (Fam.Text='') or (Imya.Text='') or (Otch.Text='') or (Kval.Text='')
  then
  showmessage('Ошибка №3.Обязательные поля не заполнены')
  else
  begin
    Form1.ADOQuery1.Close;
    Form1.ADOQuery1.SQL.Clear;
    Form1.ADOQuery1.SQL.Add('INSERT INTO Дипломы([Серия диплома],[Номер диплома],[Дата выдачи],Фамилия,Имя,Отчество,Специальность,Квалификация,Номер_приказа)');
    Form1.ADOQuery1.SQL.Add('VALUES ('+SD.text+','+ND.Text+','''+DV.text+''',');
    Form1.ADOQuery1.SQL.Add(''''+Fam.Text+''','''+Imya.Text+''','''+Otch.Text+'''');
    Form1.ADOQuery1.SQL.Add(','''+Prof.Items[Prof.ItemIndex]+''','''+Kval.text+''',');
    Form1.ADOQuery1.SQL.Add(''''+DBGrid1.DataSource.DataSet.Fields.Fields[0].AsString+''');');
    showmessage(Form1.ADOQuery1.SQL.Text);
    Form1.ADOQuery1.ExecSQL;
    showmessage('Диплом успешно зарегистрирован!');
    end
else
showmessage('Ошибка №2.Недостаточно прав доступа!');
end;

end.

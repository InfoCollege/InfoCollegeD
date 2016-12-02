unit Unit12;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB;

type
  TForm12 = class(TForm)
    Image1: TImage;
    Label2: TLabel;
    Label9: TLabel;
    FO: TEdit;
    Label1: TLabel;
    Label3: TLabel;
    FZ: TEdit;
    IO: TEdit;
    IZ: TEdit;
    OO: TEdit;
    OZ: TEdit;
    Disp: TEdit;
    KCH: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    AutoD: TButton;
    InsertTab: TButton;
    Label10: TLabel;
    InsertJZ: TADOQuery;
    procedure AutoDClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure InsertTabClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form12: TForm12;

implementation

{$R *.dfm}

uses Unit10, Unit11;

procedure TForm12.AutoDClick(Sender: TObject);
begin
Form10.Teacher.Close;
Form10.Teacher.sql.Clear;
Form10.Teacher.SQL.Add('SELECT ИД,Фамилия,Имя,Отчество');
Form10.Teacher.SQL.Add('FROM Преподаватели');
Form10.Teacher.SQL.Add('WHERE Фамилия =:k;');
Form10.Teacher.Parameters.ParamByName('k').Value:=FO.Text;
Form10.Teacher.open;
IO.Text:=Form10.DataSource1.DataSet.FindField('Имя').AsString;
OO.Text:=Form10.DataSource1.DataSet.FindField('Отчество').AsString;
Form10.Teacher.Close;
Form10.Teacher.sql.Clear;
Form10.Teacher.SQL.Add('SELECT ИД,Фамилия,Имя,Отчество');
Form10.Teacher.SQL.Add('FROM Преподаватели');
Form10.Teacher.SQL.Add('WHERE Фамилия =:a;');
Form10.Teacher.Parameters.ParamByName('a').Value:=FZ.Text;
Form10.Teacher.open;
IZ.Text:=Form10.DataSource1.DataSet.FindField('Имя').AsString;
OZ.Text:=Form10.DataSource1.DataSet.FindField('Отчество').AsString;
showmessage('Автозаполнение прошло успешно!');
end;

procedure TForm12.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form12.Hide;
Form11.show;
end;

procedure TForm12.InsertTabClick(Sender: TObject);
begin
InsertJZ.Close;
InsertJZ.SQL.Clear;
InsertJZ.SQL.Add('INSERT INTO ЖЗ(ФО,ИО,ОО,ФЗ,ИЗ,ОЗ,Дисциплина,[Кол-во часов])');
InsertJZ.SQL.Add('VALUES('''+FO.Text+''','''+IO.Text+''','''+OO.Text+''','''+FZ.text+''','''+IZ.text+''','''+OZ.Text+''','''+Disp.Text+''','''+KCH.Text+''');');
//showmessage(InsertJZ.SQL.Text);
InsertJZ.ExecSQL;
Form11.JournalZam.Close;
Form11.JournalZam.open;
end;

end.

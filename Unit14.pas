unit Unit14;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TForm14 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    DBGrid1: TDBGrid;
    KrInf: TEdit;
    Zadacha: TEdit;
    Label4: TLabel;
    Label3: TLabel;
    InFam: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    InImya: TEdit;
    InOtch: TEdit;
    IsFam: TEdit;
    IsImya: TEdit;
    IsOtch: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Dobavit: TButton;
    JP: TADOQuery;
    DataSource1: TDataSource;
    Button1: TButton;
    Inf: TADOQuery;
    DataSource2: TDataSource;
    procedure Button1Click(Sender: TObject);
    procedure DobavitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form14: TForm14;

implementation

{$R *.dfm}

uses Unit2;

procedure TForm14.Button1Click(Sender: TObject);
begin
Form14.Inf.Close;
Form14.Inf.sql.Clear;
Form14.Inf.SQL.Add('SELECT ИД_сотр,Фамилия,Имя,Отчество');
Form14.Inf.SQL.Add('FROM Сотрудники');
Form14.Inf.SQL.Add('WHERE Фамилия =:k;');
Form14.Inf.Parameters.ParamByName('k').Value:=InFam.Text;
Form14.Inf.open;
InImya.Text:=DataSource2.DataSet.FindField('Имя').AsString;
InOtch.Text:=DataSource2.DataSet.FindField('Отчество').AsString;
Inf.Close;
Inf.sql.Clear;
Inf.SQL.Add('SELECT ИД_сотр,Фамилия,Имя,Отчество');
Inf.SQL.Add('FROM Сотрудники');
Inf.SQL.Add('WHERE Фамилия =:a;');
Inf.Parameters.ParamByName('a').Value:=IsFam.Text;
Inf.open;
IsImya.Text:=DataSource2.DataSet.FindField('Имя').AsString;
IsOtch.Text:=DataSource2.DataSet.FindField('Отчество').AsString;
showmessage('Автозаполнение прошло успешно!');
end;

procedure TForm14.DobavitClick(Sender: TObject);
begin
Form14.JP.Close;
JP.SQL.clear;
JP.SQL.Add('INSERT INTO ЖП(ИФ,ИИ,ИО,КИ,З,ИФ1,ИИ1,ИО1)');
JP.SQL.Add('VALUES('''+InFam.Text+''','''+InImya.Text+''','''+InOtch.Text+''','''+KrInf.Text+''','''+Zadacha.Text+''','''+IsFam.Text+''','''+IsImya.text+''','''+IsOtch.Text+''');');
JP.ExecSQL;
JP.Close;
JP.SQL.Clear;
JP.SQL.Add('SELECT * FROM ЖП;');
JP.Open;
showmessage('Поручение зарегистрировано!');
end;

procedure TForm14.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form14.Hide;
MenuChoice.show;
end;

end.

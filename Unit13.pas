unit Unit13;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TForm13 = class(TForm)
    Image1: TImage;
    Label2: TLabel;
    Label1: TLabel;
    DBGrid1: TDBGrid;
    Period: TEdit;
    Label3: TLabel;
    Dobavit: TButton;
    Tabel: TADOQuery;
    DataSource1: TDataSource;
    Label4: TLabel;
    Fam: TEdit;
    Time: TEdit;
    Label5: TLabel;
    procedure DobavitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form13: TForm13;
   a:integer;
implementation

{$R *.dfm}

uses Unit10, Unit2;

procedure TForm13.DobavitClick(Sender: TObject);
begin
Form10.Teacher.Close;
Form10.Teacher.sql.Clear;
Form10.Teacher.SQL.Add('SELECT ИД,Фамилия,Имя,Отчество');
Form10.Teacher.SQL.Add('FROM Преподаватели');
Form10.Teacher.SQL.Add('WHERE Фамилия =:k;');
Form10.Teacher.Parameters.ParamByName('k').Value:=Fam.Text;
Form10.Teacher.open;
a:=Form10.DataSource1.DataSet.FindField('ИД').AsInteger;
Form10.Teacher.Close;
Form10.Teacher.sql.Clear;
Form10.Teacher.SQL.Add('INSERT INTO Табель');
Form10.Teacher.SQL.Add('VALUES(:b,'+Period.Text+','+Time.Text+');');
Form10.Teacher.Parameters.ParamByName('b').Value:=a;
showmessage(Form10.Teacher.SQL.Text);
Form10.Teacher.execsql;
Form13.Tabel.Close;
Form13.Tabel.Open;
showmessage('Запись добавлена!');
end;

procedure TForm13.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form13.Hide;
Form2.show;
end;

end.

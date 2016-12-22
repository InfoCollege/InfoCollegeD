unit Unit13;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TForm13 = class(TForm)
    Image1: TImage;
    L_University: TLabel;
    Label1: TLabel;
    DBGrid1: TDBGrid;
    Period: TEdit;
    Label3: TLabel;
    Dobavit: TButton;
    Tabel: TADOQuery;
    DS: TDataSource;
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
Teacher.Query_Teacher.Close;
Teacher.Query_Teacher.sql.Clear;
Teacher.Query_Teacher.SQL.Add('SELECT ��,�������,���,��������');
Teacher.Query_Teacher.SQL.Add('FROM �������������');
Teacher.Query_Teacher.SQL.Add('WHERE ������� =:k;');
Teacher.Query_Teacher.Parameters.ParamByName('k').Value:=Fam.Text;
Teacher.Query_Teacher.open;
a:=Teacher.DS.DataSet.FindField('��').AsInteger;
Teacher.Query_Teacher.Close;
Teacher.Query_Teacher.sql.Clear;
Teacher.Query_Teacher.SQL.Add('INSERT INTO ������');
Teacher.Query_Teacher.SQL.Add('VALUES(:b,'+Period.Text+','+Time.Text+');');
Teacher.Query_Teacher.Parameters.ParamByName('b').Value:=a;
//showmessage(Form10.Teacher.SQL.Text);
Teacher.Query_Teacher.execsql;
Form13.Tabel.Close;
Form13.Tabel.Open;
showmessage('������ ���������!');
end;

procedure TForm13.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form13.Hide;
MenuChoice.show;
end;

end.

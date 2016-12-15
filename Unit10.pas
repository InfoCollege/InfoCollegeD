unit Unit10;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB, Vcl.Grids, Vcl.DBGrids;

type
  TForm10 = class(TForm)
    Image1: TImage;
    Label2: TLabel;
    Label9: TLabel;
    DBGrid1: TDBGrid;
    Button1: TButton;
    Button3: TButton;
    Teacher: TADOQuery;
    DataSource1: TDataSource;
    TeacherDel: TADOQuery;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form10: TForm10;
implementation

{$R *.dfm}

uses Unit2;

procedure TForm10.Button1Click(Sender: TObject);
begin
Teacher.close;
Teacher.SQL.Clear;
Teacher.SQL.Add('INSERT INTO Преподаватели (Фамилия,Имя,Отчество)');
Teacher.SQL.Add('VALUES (NULL,NULL,NULL);');
Teacher.ExecSQL;
Teacher.close;
Teacher.SQL.Clear;
Teacher.SQL.Add('SELECT ИД,Фамилия,Имя,Отчество,КК');
Teacher.SQL.Add('FROM Преподаватели');
Teacher.SQL.Add('ORDER BY Фамилия');
Teacher.open;

end;

procedure TForm10.Button3Click(Sender: TObject);
begin
TeacherDel.close;
TeacherDel.SQL.Clear;
TeacherDel.SQL.Add('DELETE FROM Преподаватели');
TeacherDel.SQL.Add('WHERE ИД ='+inttostr(DBGrid1.Fields[0].AsInteger)+';');
TeacherDel.ExecSQL;
Teacher.Close;
Teacher.Open;
end;

procedure TForm10.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form10.Hide;
MenuChoice.show;
end;

end.

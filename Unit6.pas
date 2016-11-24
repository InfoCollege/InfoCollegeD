unit Unit6;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Imaging.jpeg, Vcl.ExtCtrls,
  Vcl.StdCtrls, Data.DB, Vcl.Grids, Vcl.DBGrids, Data.Win.ADODB;

type
  TForm6 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    DBGrid1: TDBGrid;
    Button2: TButton;
    FS: TEdit;
    Label1: TLabel;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    Button3: TButton;
    Button1: TButton;
    Label11: TLabel;
    procedure FSChange(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure DBGrid1Enter(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form6: TForm6;

implementation

{$R *.dfm}

uses Unit2, Unit8;


procedure TForm6.Button1Click(Sender: TObject);
begin
Form6.hide;;
Form8.show;
Form8.Fam.Text:=DBGrid1.DataSource.DataSet.Fields.Fields[1].AsString;
Form8.Imya.Text:=DBGrid1.DataSource.DataSet.Fields.Fields[2].AsString;
Form8.Otch.Text:=DBGrid1.DataSource.DataSet.Fields.Fields[3].AsString;
Form8.Gruppa.Text:=DBGrid1.DataSource.DataSet.Fields.Fields[4].AsString;

end;

procedure TForm6.Button2Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('INSERT INTO Студенты(Фамилия,Имя,Отчество,ИД_группы)');
ADOQuery1.SQL.Add('VALUES (NULL,NULL,NULL,0);');
ADOQuery1.ExecSQL;
ADOQuery1.Close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT ИД_студента AS [Код студента],Фамилия, Имя, Отчество, Студенты.ИД_группы as [Номер группы],Специальность');
ADOQuery1.SQL.Add('FROM Группа INNER JOIN Студенты ON Группа.ИД_группы=Студенты.ИД_группы');
ADOQuery1.SQL.Add('ORDER BY ИД_Студента;');
ADOQuery1.Open;
end;

procedure TForm6.Button3Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT ИД_студента AS [Код студента],Фамилия, Имя, Отчество, Студенты.ИД_группы as [Номер группы],Специальность');
ADOQuery1.SQL.Add('FROM Группа INNER JOIN Студенты ON Группа.ИД_группы=Студенты.ИД_группы');
ADOQuery1.SQL.Add('ORDER BY ИД_Студента;');
ADOQuery1.Open;
end;


procedure TForm6.DBGrid1Enter(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT ИД_студента AS [Код студента],Фамилия, Имя, Отчество, Студенты.ИД_группы as [Номер группы],Специальность');
ADOQuery1.SQL.Add('FROM Группа INNER JOIN Студенты ON Группа.ИД_группы=Студенты.ИД_группы');
ADOQuery1.SQL.Add('ORDER BY ИД_Студента;');
ADOQuery1.Open;
end;

procedure TForm6.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form6.Hide;
Form2.Show;
end;

procedure TForm6.FSChange(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT ИД_студента AS [Код студента],Фамилия, Имя, Отчество, Студенты.ИД_группы as [Номер группы],Специальность');
ADOQuery1.SQL.Add('FROM Группа INNER JOIN Студенты ON Группа.ИД_группы=Студенты.ИД_группы');
ADOQuery1.SQL.Add('WHERE Фамилия LIKE '''+FS.Text+'%'';');
//showmessage(ADOQuery1.SQL.Text);
ADOQuery1.Open;
end;

end.

unit Unit7;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Data.Win.ADODB, Vcl.StdCtrls, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TRating = class(TForm)
    Background: TImage;
    L_ModuleName: TLabel;
    Specialty: TListBox;
    L_Specialty: TLabel;
    Generate: TButton;
    Print: TButton;
    Query_Rating: TADOQuery;
    DS_Rating: TDataSource;
    T_Rating: TDBGrid;
    L_University: TLabel;
    procedure GenerateClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Rating: TRating;

implementation

{$R *.dfm}

procedure TRating.GenerateClick(Sender: TObject);
begin
Query_Rating.Close;
Query_Rating.SQL.Clear;
Query_Rating.SQL.Add('SELECT [Средний балл],Фамилия,Имя,Отчество,Специальность FROM ПК ');
Query_Rating.SQL.Add('WHERE Специальность=:P1');
Query_Rating.SQL.Add('ORDER BY [Средний балл] DESC;');
Query_Rating.Parameters.ParamByName('P1').Value:=Specialty.Items[Specialty.ItemIndex];
//showmessage(ADOQuery1.SQL.Text);
Query_Rating.Open;
end;

end.

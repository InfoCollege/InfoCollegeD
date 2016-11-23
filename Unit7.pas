unit Unit7;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Data.Win.ADODB, Vcl.StdCtrls, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TForm7 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    ListBox1: TListBox;
    Label7: TLabel;
    Button2: TButton;
    Button1: TButton;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    Label3: TLabel;
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form7: TForm7;

implementation

{$R *.dfm}

procedure TForm7.Button2Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQUery1.SQL.Clear;
ADOQuery1.SQL.Add('SELECT [Средний балл],Фамилия,Имя,Отчество,Специальность FROM ПК ');
ADOQuery1.SQL.Add('WHERE Специальность=:P1');
ADOQuery1.SQL.Add('ORDER BY [Средний балл] DESC;');
ADOQuery1.Parameters.ParamByName('P1').Value:=Listbox1.Items[ListBox1.ItemIndex];
//showmessage(ADOQuery1.SQL.Text);
ADOQuery1.Open;
end;

end.

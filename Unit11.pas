unit Unit11;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TForm11 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    DBGrid1: TDBGrid;
    Button1: TButton;
    JournalZam: TADOQuery;
    DataSource1: TDataSource;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form11: TForm11;

implementation

{$R *.dfm}

uses Unit12;

procedure TForm11.Button1Click(Sender: TObject);
begin
Form11.hide;
Form12.Show;
end;

end.

unit Unit16;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.StdCtrls, Vcl.Imaging.jpeg, Vcl.ExtCtrls;

type
  TRegistInfo = class(TForm)
    Background: TImage;
    L_ModuleName: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Button1: TButton;
    Label4: TLabel;
    Edit4: TEdit;
    DBGrid1: TDBGrid;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RegistInfo: TRegistInfo;

implementation

{$R *.dfm}

uses Unit2;

procedure TRegistInfo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
RegistInfo.Hide;
MenuChoice.show;
end;

end.

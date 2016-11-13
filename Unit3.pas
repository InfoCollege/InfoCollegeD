unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls;

type
  TForm3 = class(TForm)
    Image1: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation
  uses unit2;
{$R *.dfm}

procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form3.Hide;
Form2.show;
end;

end.

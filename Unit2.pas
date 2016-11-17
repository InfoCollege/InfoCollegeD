unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ComCtrls, StdCtrls, jpeg, ExtCtrls;

type
  TForm2 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    MonthCalendar1: TMonthCalendar;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N2: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    procedure N7Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N10Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation
    uses unit1,unit3, Unit4, Unit5;
{$R *.dfm}

procedure TForm2.N7Click(Sender: TObject);
begin
Form2.Hide;
Form3.show;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form2.Hide;
Form1.Show;
end;

procedure TForm2.N10Click(Sender: TObject);
begin
Form2.Hide;
Form4.show;
end;

procedure TForm2.N17Click(Sender: TObject);
begin
Form2.Hide;
Form5.show;
end;

end.

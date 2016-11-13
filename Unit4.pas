unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, DBCtrls, DB, ADODB, Grids, DBGrids;

type
  TForm4 = class(TForm)
    Button2: TButton;
    Image1: TImage;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    Edit3: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Edit7: TEdit;
    ListBox1: TListBox;
    Label9: TLabel;
    Label10: TLabel;
    DBGrid1: TDBGrid;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    procedure ListBox1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

uses Unit1;

{$R *.dfm}



procedure TForm4.ListBox1Click(Sender: TObject);
begin
if ListBox1.Selected[0]= true then
Edit7.Text:='Техник по компьютерным системам';
if ListBox1.Selected[1]= true then
Edit7.Text:='Техник-программист';
if ListBox1.Selected[2]= true then
Edit7.Text:='Техник-программист';
if ListBox1.Selected[3]= true then
Edit7.Text:='Техник по защите информации';
if ListBox1.Selected[4]= true then
Edit7.Text:='Специалист по земельно-имущественным отношениям';
if ListBox1.Selected[5]= true then
Edit7.Text:='Специалист по рекламе';
end;

procedure TForm4.Button2Click(Sender: TObject);
begin
Form1.ADOQuery1.Close;
Form1.ADOQuery1.SQL.Clear;
Form1.ADOQuery1.SQL.Add('INSERT INTO Дипломы');
Form1.ADOQuery1.SQL.Add('VALUES ('+Edit1.text+','+Edit2.Text+','+Edit3.text+',');
Form1.ADOQuery1.SQL.Add(''+Edit4.Text+','+Edit5.Text+','+Edit6.Text+'');
Form1.ADOQuery1.SQL.Add(','+Listbox1.Items[ListBox1.ItemIndex]+','+Edit7.text+',');
//Form1.ADOQuery1.SQL.Add(''+DBGrid1.DataSource.DataSet.Fields.Fields[0].Value+''');');
showmessage(Form1.ADOQuery1.SQL.Text);
end;

end.

unit Unit9;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB, Vcl.Grids, Vcl.DBGrids,ComObj;

type
  TForm9 = class(TForm)
    Label9: TLabel;
    DBGrid1: TDBGrid;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    Button1: TButton;
    Label11: TLabel;
    Image1: TImage;
    SaveDialog1: TSaveDialog;
    RadioGroup1: TRadioGroup;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form9: TForm9;
  num_rows, num_columns:integer;
implementation

{$R *.dfm}

uses Unit2;

procedure TForm9.Button1Click(Sender: TObject);
const
  wdAlignParagraphCenter = 1;
  wdAlignParagraphLeft = 0;
  wdAlignParagraphRight = 2;
   wdLineStyleSingle = 1;
    var
  table:variant;
wdApp, wdDoc, wdRng, wdTable : Variant;
  i, j, Res : Integer;
  D : TDateTime;
  Bm : TBookMark;
  Sd : TSaveDialog;
begin
num_rows:=0;
num_columns:=0;
 Sd :=SaveDialog1; //SaveDialog1 ��� ������ ���� �� �����.
  //���� ��������� ����� ������� �� ������, �� � �������� ��������� ���� �� �����,
  //� ������� ���������� ����������� ���� ����� ���������.
  if Sd.InitialDir = '' then Sd.InitialDir := ExtractFilePath( ParamStr(0) );
  //������ ������� ���������� �����.
  if not Sd.Execute then Exit;
  //���� ���� � �������� ������ ����������, �� ��������� ������ � �������������.
  if FileExists(Sd.FileName) then begin
    Res := MessageBox(0, '���� � �������� ������ ��� ����������. ������������?'
      ,'��������!', MB_YESNO + MB_ICONQUESTION + MB_APPLMODAL);
    if Res <> IDYES then Exit;
  end;
   //������� ��������� MS Word.
  try
    wdApp := CreateOleObject('Word.Application');
  except
    MessageBox(0, '�� ������� ��������� MS Word. �������� ��������.'
      ,'��������!', MB_OK + MB_ICONERROR + MB_APPLMODAL);
    Exit;
  end;
   //�� ����� ������� ������ ������� ���� MS Word.
  wdApp.Visible := True; //����� �������: wdApp.Visible := False;
  //������ ����� ��������.
  wdDoc := wdApp.Documents.Add;
  //���������� ����������� ���� MS Word, ���� wdApp.Visible := True.
  //��� ��������� ��������� � ������ ������� �������.
  wdApp.ScreenUpdating := False;

try
wdRng := wdDoc.Content; //��������, ������������ �� ���������� ���������.
wdRng.InsertAfter('����� �� ������ ��.�.�.������������(���)�');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('��������������� ������� �������������� ����������');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('�.������, �������������� ���������� �.29');
wdRng.InsertAfter(#13#10);
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
finally
end;
try
wdRng.Start := wdRng.End;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('������');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('�____�________ _______ �.                        �______________');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.Start := wdRng.End;
 wdRng.ParagraphFormat.Reset;
    wdRng.Font.Reset;
finally
end;
try
wdRng.Start := wdRng.End;
wdRng.InsertAfter('������������ �������� ��������������');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 12;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter('� ����� � ������� ������ �������� ����.');
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('����������:');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.InsertAfter('1.���������� ��������� �������������� �������� � 01.09.______ ����.');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
 wdRng.ParagraphFormat.Reset;
    wdRng.Font.Reset;
finally
end;
  try
    if not ADOQuery1.Active then ADOQuery1.Open;
    begin
    wdRng.InsertAfter(#13#10);
    //��������� ������� MS Word. ���� ������ ������� � ����� ��������.
    wdTable := wdDoc.Tables.Add(wdRng.Characters.Last, 2, ADOQuery1.Fields.Count);
    //��������� ����� �������.
    wdTable.Borders.InsideLineStyle := wdLineStyleSingle;
    wdTable.Borders.OutsideLineStyle := wdLineStyleSingle;
    //����� ���������� ���������.
    wdRng.ParagraphFormat.Reset;
    //������������ ���� ������� - �� ������ ����.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
    //���������� �����.
    wdRng := wdTable.Rows.Item(1).Range; //�������� ������ ������.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
    wdRng.Font.Size := 14;
    wdRng.Font.Bold := True;
    //���������� ������ ������ ������ - ��� ������ ������ � �������.
    //��� ���������� ��������� �����, �� ���������� ����� ������������ � ���� ������.
    wdRng := wdTable.Rows.Item(2).Range; //�������� ������ ������.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
    wdRng.Font.Size := 14;
    wdRng.Font.Bold := False;
    //���������� ����� �������
    end;
    for i := 0 to ADOQuery1.Fields.Count - 1 do
      wdTable.Cell(1, i + 1).Range.Text := ADOQuery1.Fields[i].DisplayName;
    //���������� ������ �������.
    ADOQuery1.DisableControls;
    Bm := ADOQuery1.GetBookMark;
    ADOQuery1.First;
    i:= 1;
     //������� ������ � ������� MS Word.
    while not ADOQuery1.Eof do begin
      Inc(i);
      //���� ���������, ��������� ����� ������ � ����� �������.
      if i > 2 then wdTable.Rows.Add;
      //���������� ������ � ������ ������� MS Word.
      for j := 0 to ADOQuery1.Fields.Count - 1 do
        wdTable.Cell(i, j + 1).Range.Text := ADOQuery1.Fields[j].AsString;
      ADOQuery1.Next;
    end;
    ADOQuery1.GotoBookMark(Bm);
    ADOQuery1.EnableControls;

     finally
  wdRng := wdDoc.Range.Characters.Last;;
  end;
  try
  wdRng.InsertAfter(#13#10);
  wdRng.InsertAfter(#13#10);
  wdRng.InsertAfter(#13#10);
  wdRng.InsertAfter('�������� ��������                              ����������� �.�.');
  wdRng.Font.Bold := true;
  wdRng.Font.Size := 16;
  wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
  wdRng.Start := wdRng.End;
  finally
  //��������� ����������� ���� MS Word. � ������, ���� wdApp.Visible := True.
    wdApp.ScreenUpdating := True;
  end;
  //��������� ����� ������ ��������������.
  wdApp.DisplayAlerts := False;
  try
    //������ ��������� � ����.
    wdDoc.SaveAs(FileName:=Sd.FileName);
  finally
    //�������� ����� ������ ��������������.
    wdApp.DisplayAlerts := True;
  end;

  //��������� �� ����� �������:

  //��������� ��������.
  //wdDoc.Close;
  //��������� MS Word.
  //wdApp.Quit;
 end;
procedure TForm9.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form9.Hide;
Form2.show;
end;

procedure TForm9.RadioGroup1Click(Sender: TObject);
begin
case RadioGroup1.ItemIndex of
0: begin
ADOQuery1.close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT *');
ADOQuery1.SQL.Add('FROM �������������');
ADOQuery1.SQL.Add('ORDER BY �������;');
ADOQuery1.open;
DBGrid1.ReadOnly:=false;
end;
1:  begin
ADOQuery1.close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT �������,���,��������,[������� ��������]');
ADOQuery1.SQL.Add('FROM �������������');
ADOQuery1.SQL.Add('ORDER BY �������;');
ADOQuery1.open;
DBGrid1.ReadOnly:=true;
end;
end;
end;
end.

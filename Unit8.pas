unit Unit8;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
 ComObj, Vcl.Imaging.jpeg;

type
  TForm8 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    Label3: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Fam: TEdit;
    Imya: TEdit;
    Otch: TEdit;
    Label5: TLabel;
    Gruppa: TEdit;
    Label6: TLabel;
    God: TEdit;
    Label7: TLabel;
    Treb: TEdit;
    Button2: TButton;
    Label8: TLabel;
    Label10: TLabel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    FIOruk: TEdit;
    DolRuk: TEdit;
    Button1: TButton;
    SaveDialog1: TSaveDialog;
    RadioGroup1: TRadioGroup;
    procedure RadioButton1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form8: TForm8;
  a:string;

implementation

{$R *.dfm}

uses Unit2, Unit6, Unit1;

procedure TForm8.Button1Click(Sender: TObject);
 const
  wdAlignParagraphCenter = 1;
  wdAlignParagraphLeft = 0;
  wdAlignParagraphRight = 2;
  var
  wdApp, wdDoc, wdRng : Variant;
  Res : Integer;
  Sd : TSaveDialog;
begin
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
wdRng.InsertAfter('����� �� "����� ��.�.�.������������(���)');
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
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('������� �');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('������  '+Fam.Text+' '+Imya.Text+' '+Otch.Text+'');
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter('� ���, ��� ��(���) �������� (��������) � '+God.Text+' ������� ���� ��������� ����� ����� ��������   ');
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('������������ ���������������� ���������� ���������������� ���������� ������� ����������� "���������� ��������������� ����������� ���������� � ���������� ����� �.�.������������ (������ ������� �����������)');
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('������� ������  '+Gruppa.Text+' �� ��������� ���  '+a+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('������� ���� ��� �������������� �  '+Treb.Text+'.');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter(''+DolRuk.Text+'                                          '+FioRuk.Text+'');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.Font.Bold := True;
wdRng.Font.Size := 15;
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


procedure TForm8.Button2Click(Sender: TObject);
begin
Form8.hide;
Form6.show;
end;

procedure TForm8.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form8.Hide;
Form2.show;
end;

procedure TForm8.RadioButton1Click(Sender: TObject);
begin
FIORuk.Text:='����������� �.�.';
DolRuk.Text:='�������� ��������';
end;

procedure TForm8.RadioButton2Click(Sender: TObject);
begin
FIORuk.text:=Form1.DataSource1.DataSet.FindField('����_���').AsString;
DolRuk.text:=Form1.DataSource1.DataSet.FindField('������������').AsString;
end;



procedure TForm8.RadioGroup1Click(Sender: TObject);
begin
case RadioGroup1.ItemIndex of
0: a:=Form6.DBGrid1.DataSource.DataSet.FindField('�������������').AsString;
1: a:='';
end;

end;

end.

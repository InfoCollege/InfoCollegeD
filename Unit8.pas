unit Unit8;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
 ComObj, Vcl.Imaging.jpeg;

type
  TMenuGenerate = class(TForm)
    Background: TImage;
    L_ModuleName: TLabel;
    L_University: TLabel;
    L_Surname: TLabel;
    L_Name: TLabel;
    L_MiddleName: TLabel;
    Surname: TEdit;
    Name: TEdit;
    MiddleName: TEdit;
    L_Group: TLabel;
    Group: TEdit;
    L_Date: TLabel;
    Date: TEdit;
    L_Destination: TLabel;
    Destination: TEdit;
    handbook: TButton;
    L_director: TLabel;
    L_Position: TLabel;
    Director1: TRadioButton;
    Director2: TRadioButton;
    director: TEdit;
    position: TEdit;
    Generate: TButton;
    SaveDialog1: TSaveDialog;
    AddInfo: TRadioGroup;
    procedure Director1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GenerateClick(Sender: TObject);
    procedure handbookClick(Sender: TObject);
    procedure Director2Click(Sender: TObject);
    procedure AddInfoClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MenuGenerate: TMenuGenerate;
  a:string;

implementation

{$R *.dfm}

uses Unit2, Unit6, Unit1;

procedure TMenuGenerate.GenerateClick(Sender: TObject);
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
wdRng.InsertAfter(''+MainForm.INFO.Fields[0].AsString+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(''+MainForm.INFO.Fields[2].AsString+'');
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
 wdRng.InsertAfter('������  '+Surname.Text+' '+Name.Text+' '+Middlename.Text+'');
 wdRng.InsertAfter(#13#10);
 wdRng.Font.Bold := False;
 wdRng.Font.Size := 14;
 wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
 wdRng.InsertAfter('� ���, ��� ��(���) �������� (��������) � '+Date.Text+' ������� ���� ��������� ����� ����� ��������   ');
 wdRng.Font.Bold := False;
 wdRng.Font.Size := 14;
 wdRng.Start := wdRng.End;
 wdRng.ParagraphFormat.Reset;
 wdRng.Font.Reset;
 wdRng.InsertAfter(''+MainForm.INFO.Fields[0].AsString+'');
 wdRng.InsertAfter(#13#10);
 wdRng.Font.Bold := False;
 wdRng.Font.Size := 14;
 wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
 wdRng.Start := wdRng.End;
 wdRng.ParagraphFormat.Reset;
 wdRng.Font.Reset;
 wdRng.InsertAfter('������� ������  '+Group.Text+' �� ��������� ���  '+a+'');
 wdRng.InsertAfter(#13#10);
 wdRng.InsertAfter('������� ���� ��� �������������� �  '+Destination.Text+'.');
 wdRng.InsertAfter(#13#10);
 wdRng.InsertAfter(#13#10);
 wdRng.Font.Bold := False;
 wdRng.Font.Size := 14;
 wdRng.Start := wdRng.End;
 wdRng.ParagraphFormat.Reset;
 wdRng.Font.Reset;
 wdRng.InsertAfter(''+Position.Text+'                                          '+Director.Text+'');
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


procedure TMenuGenerate.handbookClick(Sender: TObject);
begin
Generate.hide;
RegisterStudent.show;
end;

procedure TMenuGenerate.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Generate.Hide;
MenuChoice.show;
end;

procedure TMenuGenerate.Director1Click(Sender: TObject);
begin
Director.Text:=''+MainForm.INFO.Fields[3].AsString+'';
Position.Text:='�������� ��������';
end;

procedure TMenuGenerate.Director2Click(Sender: TObject);
begin
Director.text:=MainForm.AuthDS1.DataSet.FindField('����_���').AsString;
Position.text:=MainForm.AuthDS1.DataSet.FindField('������������').AsString;
end;



procedure TMenuGenerate.AddInfoClick(Sender: TObject);
begin
case AddInfo.ItemIndex of
0: a:=RegisterStudent.T_RegisterStudent.DataSource.DataSet.FindField('�������������').AsString;
1: a:='';
end;

end;

end.

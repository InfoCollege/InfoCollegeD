unit Unit5;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB,ComObj;

type
  TAdmissionCommittee = class(TForm)
    Background: TImage;
    L_ModuleName: TLabel;
    L_Surname: TLabel;
    L_Name: TLabel;
    L_MiddleName: TLabel;
    L_Information1: TLabel;
    L_Information2: TLabel;
    L_NameSchool: TLabel;
    L_Date: TLabel;
    L_Information3: TLabel;
    L_Excelent: TLabel;
    L_Good: TLabel;
    L_Satisfactory: TLabel;
    L_Average: TLabel;
    L_Information5: TLabel;
    L_Speciality: TLabel;
    Specialty: TListBox;
    L_Information4: TLabel;
    L_Passport: TLabel;
    L_DatePassport: TLabel;
    L_PassportIssued: TLabel;
    L_ID: TLabel;
    L_PhoneHouse: TLabel;
    L_Information6: TLabel;
    L_Phone: TLabel;
    L_Information7: TLabel;
    L_FullName: TLabel;
    Surname: TEdit;
    Name: TEdit;
    MiddleName: TEdit;
    NameSchool: TEdit;
    Date: TEdit;
    Passport: TEdit;
    DatePassport: TEdit;
    PassportIssued: TEdit;
    ID: TEdit;
    Excelent: TEdit;
    Good: TEdit;
    Satisfactory: TEdit;
    Average: TEdit;
    PhoneHouse: TEdit;
    Phone: TEdit;
    Fullname: TEdit;
    Register: TButton;
    �ounting: TButton;
    PrintCard: TButton;
    SaveDialog1: TSaveDialog;
    Query_Register: TADOQuery;
    L_University: TLabel;
    procedure �ountingClick(Sender: TObject);
    procedure RegisterClick(Sender: TObject);
    procedure PrintCardClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AdmissionCommittee: TAdmissionCommittee;
  sr:integer;
  sb:real;
  pred:integer;
implementation

{$R *.dfm}


procedure TAdmissionCommittee.�ountingClick(Sender: TObject);
begin
if (Excelent.Text='') OR (Good.Text='') OR (Satisfactory.Text='') then
 showmessage('������ �3. ������������ ���� �� ���������!')
 else
 begin
 pred:=strtoint(Excelent.Text)+strtoint(Good.Text)+strtoint(Satisfactory.Text);
 sr:=(5*strtoint(Excelent.Text))+(4*strtoint(Good.Text))+(3*strtoint(Satisfactory.Text));
 sb:=sr/pred;
 Average.Text:=FloatToStr(sb);
 end
 end;

procedure TAdmissionCommittee.RegisterClick(Sender: TObject);
begin
Query_Register.Close;
Query_Register.SQL.Clear;
Query_Register.SQL.Add('INSERT INTO ��(�������,���,��������,��������������,�������,�������,����������,�����,���,[������� ����],�������������,�������,[�������(2)],����������)');
Query_Register.SQL.Add('VALUES('''+Surname.Text+''','''+Name.Text+''','''+MiddleName.Text+''','''+NameSchool.Text+''','''+Date.Text+''','''+Passport.Text+''','''+DatePassport.Text+''','''+PassportIssued.Text+''','''+ID.Text+''','''+Average.Text+'''');
Query_Register.SQL.Add(','''+Specialty.Items[Specialty.ItemIndex]+''','''+PhoneHouse.Text+''','''+Phone.Text+''','''+FullName.Text+''');');
//showmessage(Query_Register.SQL.Text);
Query_Register.ExecSQL;
showmessage('���������� ������ � ������');
end;

procedure TAdmissionCommittee.PrintCardClick(Sender: TObject);
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
wdRng.InsertBefore('������ �������������� ������� "�������������� �������"');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('�������� �������� ');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore(' '+Surname.text+'  '+Name.text+'  '+MiddleName.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('�����������');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore(''+NameSchool.text+',������� �  '+Date.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('���������� ������');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('����� � ����� ��������:'+Passport.text+' ���� ������:'+DatePassport.text+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('�����:'+PassportIssued.text+' ��� �������������:'+ID.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('������� ����:'+Average.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('�������������'+Specialty.Items[Specialty.ItemIndex]+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('��������');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('��������� �.:'+PhoneHouse.text+' ���������:'+Phone.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('�������������� ������ �����������:');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('��������� ��:'+FullName.text+'_______________');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('����������:'+Surname.text+''+Copy(Name.text,1,1)+'.'+Copy(MiddleName.text,1,1)+'.''_______________');
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
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

end.

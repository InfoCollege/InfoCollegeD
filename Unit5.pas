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
    Сounting: TButton;
    PrintCard: TButton;
    SaveDialog1: TSaveDialog;
    Query_Register: TADOQuery;
    L_University: TLabel;
    procedure СountingClick(Sender: TObject);
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


procedure TAdmissionCommittee.СountingClick(Sender: TObject);
begin
if (Excelent.Text='') OR (Good.Text='') OR (Satisfactory.Text='') then
 showmessage('Ошибка №3. Обязательные поля не заполенны!')
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
Query_Register.SQL.Add('INSERT INTO ПК(Фамилия,Имя,Отчество,НаименованиеОУ,Окончил,Паспорт,Датавыдачи,Выдан,Код,[Средний балл],Специальность,Телефон,[Телефон(2)],Примечание)');
Query_Register.SQL.Add('VALUES('''+Surname.Text+''','''+Name.Text+''','''+MiddleName.Text+''','''+NameSchool.Text+''','''+Date.Text+''','''+Passport.Text+''','''+DatePassport.Text+''','''+PassportIssued.Text+''','''+ID.Text+''','''+Average.Text+'''');
Query_Register.SQL.Add(','''+Specialty.Items[Specialty.ItemIndex]+''','''+PhoneHouse.Text+''','''+Phone.Text+''','''+FullName.Text+''');');
//showmessage(Query_Register.SQL.Text);
Query_Register.ExecSQL;
showmessage('Абитуриент внесен в реестр');
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
  Sd :=SaveDialog1; //SaveDialog1 уже должен быть на форме.
  //Если начальная папка диалога не задана, то в качестве начальной берём ту папку,
  //в которой расположен исполняемый файл нашей программы.
  if Sd.InitialDir = '' then Sd.InitialDir := ExtractFilePath( ParamStr(0) );
  //Запуск диалога сохранения файла.
  if not Sd.Execute then Exit;
  //Если файл с заданным именем существует, то запускаем диалог с пользователем.
  if FileExists(Sd.FileName) then begin
    Res := MessageBox(0, 'Файл с заданным именем уже существует. Перезаписать?'
      ,'Внимание!', MB_YESNO + MB_ICONQUESTION + MB_APPLMODAL);
    if Res <> IDYES then Exit;
  end;

  //Попытка запустить MS Word.
  try
    wdApp := CreateOleObject('Word.Application');
  except
    MessageBox(0, 'Не удалось запустить MS Word. Действие отменено.'
      ,'Внимание!', MB_OK + MB_ICONERROR + MB_APPLMODAL);
    Exit;
  end;

  //На время отладки делаем видимым окно MS Word.
  wdApp.Visible := True; //После отладки: wdApp.Visible := False;
  //Создаём новый документ.
  wdDoc := wdApp.Documents.Add;
  //Отключение перерисовки окна MS Word, если wdApp.Visible := True.
  //Для ускорения обработки в случае больших текстов.
  wdApp.ScreenUpdating := False;
try
wdRng := wdDoc.Content; //Диапазон, охватывающий всё содержимое документа.
wdRng.InsertBefore('ЕДИНАЯ ИНФОРМАЦИОННАЯ СИСТЕМА "ИНФОРМАЦИОННЫЙ КОЛЛЕДЖ"');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('ПРИЕМНАЯ КОМИССИЯ ');
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
wdRng.InsertBefore('Образование');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore(''+NameSchool.text+',окончил в  '+Date.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Паспортные данные');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Серия и номер паспорта:'+Passport.text+' Дата выдачи:'+DatePassport.text+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Выдан:'+PassportIssued.text+' Код подразделения:'+ID.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Средний балл:'+Average.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Специальность'+Specialty.Items[Specialty.ItemIndex]+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Контакты');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Городской т.:'+PhoneHouse.text+' Мобильный:'+Phone.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Представленные данные подтверждаю:');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Секретарь ПК:'+FullName.text+'_______________');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Абитуриент:'+Surname.text+''+Copy(Name.text,1,1)+'.'+Copy(MiddleName.text,1,1)+'.''_______________');
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
  finally
    //Включение перерисовки окна MS Word. В случае, если wdApp.Visible := True.
    wdApp.ScreenUpdating := True;
  end;

  //Отключаем режим показа предупреждений.
  wdApp.DisplayAlerts := False;
  try
    //Запись документа в файл.
    wdDoc.SaveAs(FileName:=Sd.FileName);
  finally
    //Включаем режим показа предупреждений.
    wdApp.DisplayAlerts := True;
  end;

  //Отключено на время отладки:

  //Закрываем документ.
  //wdDoc.Close;
  //Закрываем MS Word.
  //wdApp.Quit;
end;

end.

unit Unit5;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, Data.DB, Data.Win.ADODB,ComObj;

type
  TForm5 = class(TForm)
    Image1: TImage;
    Label9: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    ListBox1: TListBox;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit15: TEdit;
    Edit16: TEdit;
    Button2: TButton;
    Button1: TButton;
    ADOQuery1: TADOQuery;
    Button3: TButton;
    SaveDialog1: TSaveDialog;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form5: TForm5;
  sr:integer;
  sb:real;
  pred:integer;
implementation

{$R *.dfm}


procedure TForm5.Button1Click(Sender: TObject);
begin
pred:=strtoint(Edit10.Text)+strtoint(Edit11.Text)+strtoint(Edit12.Text);
 sr:=(5*strtoint(Edit10.Text))+(4*strtoint(Edit11.Text))+(3*strtoint(Edit12.Text));
 sb:=sr/pred;
 Edit13.Text:=FloatToStr(sb);
end;

procedure TForm5.Button2Click(Sender: TObject);
begin
ADOQuery1.Close;
ADOQuery1.SQL.Clear;
ADOQuery1.SQL.Add('INSERT INTO ПК(Фамилия,Имя,Отчество,НаименованиеОУ,Окончил,Паспорт,Датавыдачи,Выдан,Код,[Средний балл],Специальность,Телефон,[Телефон(2)],Примечание)');
ADOQuery1.SQL.Add('VALUES('''+Edit1.Text+''','''+Edit2.Text+''','''+Edit3.Text+''','''+Edit4.Text+''','''+Edit5.Text+''','''+Edit6.Text+''','''+Edit7.Text+''','''+Edit8.Text+''','''+Edit9.Text+''','''+Edit13.Text+'''');
ADOQuery1.SQL.Add(','''+Listbox1.Items[ListBox1.ItemIndex]+''','''+Edit14.Text+''','''+Edit15.Text+''','''+Edit16.Text+''');');
//showmessage(ADOQuery1.SQL.Text);
ADOQuery1.ExecSQL;
showmessage('Абитуриент внесен в реестр');
end;

procedure TForm5.Button3Click(Sender: TObject);
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
wdRng.InsertBefore('ПРИЕМНАЯ КОМИССИЯ УКИТ');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore(' '+Edit1.text+'  '+Edit2.text+'  '+Edit3.text+'');
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
wdRng.InsertBefore(''+Edit4.text+',окончил в  '+Edit5.text+'');
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
wdRng.InsertBefore('Серия и номер паспорта:'+Edit6.text+' Дата выдачи:'+Edit7.text+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Выдан:'+Edit8.text+' Код подразделения:'+Edit9.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Средний балл:'+Edit13.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Специальность'+Listbox1.Items[ListBox1.ItemIndex]+'');
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
wdRng.InsertBefore('Городской т.:'+Edit14.text+' Мобильный:'+Edit15.text+'');
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertBefore('Представленные данные подтверждаю:');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Секретарь ПК:'+Edit16.text+'_______________');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Абитуриент:'+Edit1.text+''+Copy(Edit2.text,1,1)+'.'+Copy(Edit3.text,1,1)+'.''_______________');
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

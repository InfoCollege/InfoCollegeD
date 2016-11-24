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
wdRng.InsertAfter('ФГБОУ ВО "МГУТУ им.К.Г.Разумовского(ПКУ)');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Университетский колледж информационных технологий');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('г.Москва, Костомаровская набережная д.29');
wdRng.InsertAfter(#13#10);
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('СПРАВКА №');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('Выдана  '+Fam.Text+' '+Imya.Text+' '+Otch.Text+'');
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter('В том, что он(она) является (являлась) в '+God.Text+' учебном году студентом очной формы обучения   ');
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('Федерального государственного бюджетного образовательного учреждения высшего образования "Московский государственный университет технологий и управления имени К.Г.Разумовского (Первый казачий университет)');
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('Учебная группа  '+Gruppa.Text+' по программе СПО  '+a+'');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Справка дана для предоставления в  '+Treb.Text+'.');
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
FIORuk.Text:='Александров Р.В.';
DolRuk.Text:='Директор колледжа';
end;

procedure TForm8.RadioButton2Click(Sender: TObject);
begin
FIORuk.text:=Form1.DataSource1.DataSet.FindField('Сокр_имя').AsString;
DolRuk.text:=Form1.DataSource1.DataSet.FindField('Наименование').AsString;
end;



procedure TForm8.RadioGroup1Click(Sender: TObject);
begin
case RadioGroup1.ItemIndex of
0: a:=Form6.DBGrid1.DataSource.DataSet.FindField('Специальность').AsString;
1: a:='';
end;

end;

end.

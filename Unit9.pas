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
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
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
    var
  wdApp, wdDoc, wdRng : Variant;
  Res : Integer;
  Sd : TSaveDialog;
  table:variant;
begin
num_rows:=0;
num_columns:=0;
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
wdRng.InsertAfter('ПРИКАЗ');
wdRng.InsertAfter('"____"________ _______ г.                                          №______________');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('Утверждение нагрузки преподавателей');
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('В связи с началом нового учебного года.');
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('ПРИКАЗЫВАЮ:');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
wdRng.InsertAfter(#13#10);
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter('1.Установить следующую педагогическую нагрузку с 01.09.______ года.');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
 if num_rows = num_columns then
  begin
    //num_colums:=ADOQuery1.RecordCount;
    //wdRng.InsertAfter(DataSource1.DataSet.FindField('Фамилия');
  end;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);

wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.Start := wdRng.End;
wdRng.ParagraphFormat.Reset;
wdRng.Font.Reset;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('Директор колледжа                              Александров Р.В.');
wdRng.Font.Bold := true;
wdRng.Font.Size := 16;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
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
procedure TForm9.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form9.Hide;
Form2.show;
end;

end.

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
wdRng.InsertAfter('ФГБОУ ВО «МГУТУ им.К.Г.Разумовского(ПКУ)»');
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
finally
end;
try
wdRng.Start := wdRng.End;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('ПРИКАЗ');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('«____»________ _______ г.                        №______________');
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
wdRng.InsertAfter('«Утверждение нагрузки преподавателей»');
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := False;
wdRng.Font.Size := 12;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter('В связи с началом нового учебного года.');
wdRng.Font.Bold := False;
wdRng.Font.Size := 14;
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.InsertAfter('ПРИКАЗЫВАЮ:');
wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
wdRng.InsertAfter(#13#10);
wdRng.Font.Name := 'Times New Roman';
wdRng.Font.Bold := True;
wdRng.Font.Size := 18;
wdRng.InsertAfter('1.Установить следующую педагогическую нагрузку с 01.09.______ года.');
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
    //Добавляем таблицу MS Word. Пока создаём таблицу с двумя строками.
    wdTable := wdDoc.Tables.Add(wdRng.Characters.Last, 2, ADOQuery1.Fields.Count);
    //Параметры линий таблицы.
    wdTable.Borders.InsideLineStyle := wdLineStyleSingle;
    wdTable.Borders.OutsideLineStyle := wdLineStyleSingle;
    //Сброс параметров параграфа.
    wdRng.ParagraphFormat.Reset;
    //Выравнивание всей таблицы - по левому краю.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
    //Оформление шапки.
    wdRng := wdTable.Rows.Item(1).Range; //Диапазон первой строки.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
    wdRng.Font.Size := 14;
    wdRng.Font.Bold := True;
    //Оформление первой строки данных - это вторая строка в таблице.
    //При добавлении следующих строк, их оформление будет копироваться с этой строки.
    wdRng := wdTable.Rows.Item(2).Range; //Диапазон второй строки.
    wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
    wdRng.Font.Size := 14;
    wdRng.Font.Bold := False;
    //Записываем шапку таблицы
    end;
    for i := 0 to ADOQuery1.Fields.Count - 1 do
      wdTable.Cell(1, i + 1).Range.Text := ADOQuery1.Fields[i].DisplayName;
    //Записываем данные таблицы.
    ADOQuery1.DisableControls;
    Bm := ADOQuery1.GetBookMark;
    ADOQuery1.First;
    i:= 1;
     //Текущая строка в таблице MS Word.
    while not ADOQuery1.Eof do begin
      Inc(i);
      //Если требуется, добавляем новую строку в конец таблицы.
      if i > 2 then wdTable.Rows.Add;
      //Записываем данные в строку таблицы MS Word.
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
  wdRng.InsertAfter('Директор колледжа                              Александров Р.В.');
  wdRng.Font.Bold := true;
  wdRng.Font.Size := 16;
  wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
  wdRng.Start := wdRng.End;
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

procedure TForm9.RadioGroup1Click(Sender: TObject);
begin
case RadioGroup1.ItemIndex of
0: begin
ADOQuery1.close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT *');
ADOQuery1.SQL.Add('FROM Преподаватели');
ADOQuery1.SQL.Add('ORDER BY Фамилия;');
ADOQuery1.open;
DBGrid1.ReadOnly:=false;
end;
1:  begin
ADOQuery1.close;
ADOQuery1.SQL.clear;
ADOQuery1.SQL.Add('SELECT Фамилия,Имя,Отчество,[Годовая нагрузка]');
ADOQuery1.SQL.Add('FROM Преподаватели');
ADOQuery1.SQL.Add('ORDER BY Фамилия;');
ADOQuery1.open;
DBGrid1.ReadOnly:=true;
end;
end;
end;
end.

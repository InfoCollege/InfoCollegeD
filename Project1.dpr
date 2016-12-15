program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {MainForm},
  Unit2 in 'Unit2.pas' {MenuChoice},
  Unit3 in 'Unit3.pas' {Information},
  Unit4 in 'Unit4.pas' {RegisterDiplomas},
  Unit5 in 'Unit5.pas' {AdmissionCommittee},
  Unit6 in 'Unit6.pas' {RegisterStudent},
  Unit7 in 'Unit7.pas' {Rating},
  Unit8 in 'Unit8.pas' {Form8},
  Unit9 in 'Unit9.pas' {Form9},
  Unit10 in 'Unit10.pas' {Form10},
  Unit11 in 'Unit11.pas' {Form11},
  Unit12 in 'Unit12.pas' {Form12},
  Unit13 in 'Unit13.pas' {Form13},
  Unit14 in 'Unit14.pas' {Form14};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'ЕИС "Информационный колледж"';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TMenuChoice, MenuChoice);
  Application.CreateForm(TInformation, Information);
  Application.CreateForm(TRegisterDiplomas, RegisterDiplomas);
  Application.CreateForm(TAdmissionCommittee, AdmissionCommittee);
  Application.CreateForm(TRegisterStudent, RegisterStudent);
  Application.CreateForm(TRating, Rating);
  Application.CreateForm(TForm8, Form8);
  Application.CreateForm(TForm9, Form9);
  Application.CreateForm(TForm10, Form10);
  Application.CreateForm(TForm11, Form11);
  Application.CreateForm(TForm12, Form12);
  Application.CreateForm(TForm13, Form13);
  Application.CreateForm(TForm14, Form14);
  Application.Run;
end.

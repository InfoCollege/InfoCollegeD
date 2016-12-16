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
  Unit8 in 'Unit8.pas' {MenuGenerate},
  Unit9 in 'Unit9.pas' {workload},
  Unit10 in 'Unit10.pas' {Teacher},
  Unit11 in 'Unit11.pas' {JournalReplacment},
  Unit12 in 'Unit12.pas' {Form12},
  Unit13 in 'Unit13.pas' {Form13},
  Unit14 in 'Unit14.pas' {TaskBook},
  Unit15 in 'Unit15.pas' {MethodicalCabinet};

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
  Application.CreateForm(TMenuGenerate, MenuGenerate);
  Application.CreateForm(Tworkload, workload);
  Application.CreateForm(TTeacher, Teacher);
  Application.CreateForm(TJournalReplacment, JournalReplacment);
  Application.CreateForm(TForm12, Form12);
  Application.CreateForm(TForm13, Form13);
  Application.CreateForm(TTaskBook, TaskBook);
  Application.CreateForm(TMethodicalCabinet, MethodicalCabinet);
  Application.Run;
end.

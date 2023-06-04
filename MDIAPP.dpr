program Mdiapp;

uses
  Forms,
  main in 'main.pas' {MainForm},
  About in 'About.pas' {AboutBox},
  DataUnit in 'DataUnit.pas' {DataMod: TDataModule},
  RepCodeForMonths in 'RepCodeForMonths.pas' {RepCodetForMonthsForm},
  RepStatForRecWin in 'RepStatForRecWin.pas' {RepStatForRecForm},
  MonthWin in 'MonthWin.pas' {MonthDLG},
  MonthGridType in 'MonthGridType.pas',
  RepStatForMonths in 'RepStatForMonths.pas' {RepStatForMonthsForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TDataMod, DataMod);
  Application.Run;
end.

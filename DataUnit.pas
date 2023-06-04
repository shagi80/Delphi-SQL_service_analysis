unit DataUnit;

interface

uses
  SysUtils, Classes, Messages, ImgList, Controls, DB, ADODB, IniFiles,
  Forms, Dialogs, DateUtils;

const
  WM_CLOSECHILD=WM_USER+1;

const
  tmDay   = 101;
  tmWeek  = 107;
  tmHalf  = 115;
  tmMonth = 130;

type
  //период дат
  TDatePeriod = record
    first,last : TDateTime;
  end;
  //месяц года, разбитый по периодам дат
  TMonthPeriod = record
    month, year : word;
    dates : array of TDatePeriod;
  end;
  //период планирования
  TTimePeriod = class(TList)
  private
    function   GetMonth(i: integer): TMonthPeriod;
    function   GetPeriod(i: integer): TDatePeriod;
    function   GetPeriodCount : integer;
  public
    procedure  Clear; override;
    function   SetDates(firstdate, lastdate : TDateTime; mode : integer):integer;
    property   Month[i: integer]: TMonthPeriod read GetMonth;
    property   PeriodsCount : integer read GetPeriodCount;
    property   Period[i: integer]: TDatePeriod read GetPeriod;
  end;

  TDataMod = class(TDataModule)
    Connection: TADOConnection;
    ButtonImages: TImageList;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DataMod: TDataMod;

implementation

{$R *.dfm}


procedure TDataMod.DataModuleCreate(Sender: TObject);
var
  myini : TIniFile;
  path  : string;
begin
  if Connection.Connected then Connection.Connected:=false;
  path:=ExtractFilePath(application.ExeName);
  path:=path+'myini.ini';
  myini:=TIniFile.Create(path);
  Connection.ConnectionString:=myini.ReadString('MAIN','CONSTR','');
  try
    Connection.Connected:=true;
  except
    MessageDlg('Ошибка подключения к БД !',mtError,[mbOk],0);
    Halt(0);
  end;
end;

//----------------------- Класс периодов дат -----------------------------------

procedure TTimePeriod.Clear;
var
  i,j : integer;
begin
  for I := 0 to self.Count - 1 do begin
    if (self.Items[i]<>nil) then begin
      TMonthPeriod(self.Items[i]^).month:=0;
      TMonthPeriod(self.Items[i]^).year:=0;
      SetLength(TMonthPeriod(self.Items[i]^).dates,0);
      FreeMem(self.Items[i],sizeof(TMonthPeriod));
    end;
  end;
  inherited;
end;

function  TTimePeriod.GetMonth(i: integer): TMonthPeriod;
begin
  result:=TMonthPeriod(self.Items[i]^);
end;

function  TTimePeriod.GetPeriodCount : integer;
var
  j,sum : integer;
begin
  sum:=0;
  for j := 0 to self.Count-1 do sum:=sum+high(self.Month[j].dates)+1;
  result:=sum;
end;

function  TTimePeriod.GetPeriod(i: Integer): TDatePeriod;
var
  j,k,sum : integer;
begin
  j:=0; sum:=0; k:=0;
  while (sum<>i)do begin
    inc(k);
    inc(sum);
    if k>high(self.Month[j].dates) then begin
      inc(j);
      k:=0;
    end;
  end;
  if (sum>=i)then result:=self.Month[j].dates[k];
end;

function  TTimePeriod.SetDates(firstdate, lastdate : TDateTime; mode : integer):integer;
var
  mn,dt     : TDateTime;
  month     : ^ TMonthPeriod;
  nextmonth : boolean;
begin
  if self.Count>0 then self.Clear;
  mn:=StartOfTheMonth(firstdate);
  dt:=firstdate;
  repeat
    nextmonth:=false;
    new(month);
    month^.year:=YearOf(mn);
    month^.month:=MonthOf(mn);
    SetLength(month.dates,0);
    repeat
      SetLength(month.dates,high(month.dates)+2);
      month.dates[high(month.dates)].first:=dt;
      case mode of
        tmMonth : dt:=EndOfTheMonth(mn);
        tmHalf  : if DayOf(dt)<=15 then dt:=IncDay(mn,14) else dt:=EndOfTheMonth(mn);
        tmWeek  : dt:=EndOfTheWeek(dt);
        tmDay   : ;
        else dt:=IncDay(dt,mode-1);
      end;
      month.dates[high(month.dates)].last:=dt;
      dt:=IncDay(dt,1);
      nextmonth:=(DaysBetween(dt,EndOfTheMonth(mn))<3);
      if not nextmonth then nextmonth:=(dt>EndOfTheMonth(mn));
    until (nextmonth)or(dt>lastdate);
    if dt>lastdate then month.dates[high(month.dates)].last:=lastdate;
    self.Add(month);
    mn:=IncMonth(mn,1);
  until dt>lastdate;
  result:=self.Count;
end;




end.

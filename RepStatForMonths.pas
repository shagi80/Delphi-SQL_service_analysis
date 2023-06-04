unit RepStatForMonths;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ComCtrls, CheckLst, DB, ADODB, Grids,
  DBGrids, DBCtrls, frxClass, frxExportPDF, frxExportXLS,
  frxExportODF, frxDBSet, MonthGridType;

type
  TRepStatForMonthsForm = class(TForm)
    TopPn: TPanel;
    CaptionLB: TLabel;
    MainPn: TPanel;
    SB: TScrollBox;
    FilterPn: TPanel;
    BottomBevel: TBevel;
    Panel4: TPanel;
    WorkTypeCheck: TCheckBox;
    WorkTypeCB: TComboBox;
    ModelPn: TPanel;
    ModelCheck: TCheckBox;
    ModelCB: TCheckListBox;
    ClientPn: TPanel;
    ClientDataCheck: TCheckBox;
    ClientDataPn: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    ClearClientBtn: TSpeedButton;
    ClientAddrBtn: TSpeedButton;
    ClientTelBtn: TSpeedButton;
    ClientEd: TEdit;
    AddrEd: TEdit;
    TelEd: TEdit;
    Panel5: TPanel;
    ClearSNEDBTN: TSpeedButton;
    SNCheck: TCheckBox;
    SNEd: TEdit;
    Panel6: TPanel;
    CostCheck: TCheckBox;
    CostDataPn: TPanel;
    Label6: TLabel;
    Label7: TLabel;
    Label5: TLabel;
    PartCostMin: TEdit;
    PartCostMax: TEdit;
    MovCostMin: TEdit;
    MovCostMax: TEdit;
    WorkCostMin: TEdit;
    WorkCostMax: TEdit;
    WorkForClientPn: TPanel;
    WorkForClientCheck: TCheckBox;
    WorkForClientED: TEdit;
    SenterPN: TPanel;
    SenterCheck: TCheckBox;
    SenterCB: TCheckListBox;
    Panel8: TPanel;
    ClearWorkTimeBtn: TSpeedButton;
    WorkTimeCheck: TCheckBox;
    WorkTimeED: TEdit;
    WorkTimeCB: TComboBox;
    Panel9: TPanel;
    ClearPartBtn: TSpeedButton;
    PartCheck: TCheckBox;
    PartED: TEdit;
    Panel10: TPanel;
    ClearProblemNoteBtn: TSpeedButton;
    ProblemNoteCheck: TCheckBox;
    ProblemNoteED: TEdit;
    Panel11: TPanel;
    ClearWorkNoteBtn: TSpeedButton;
    WorkNoteCheck: TCheckBox;
    WorkNoteED: TEdit;
    Panel12: TPanel;
    WorkCodeCheck: TCheckBox;
    WorkCodeCB: TCheckListBox;
    SetBtn: TSpeedButton;
    Panel1: TPanel;
    BottomPn: TPanel;
    PrepareBtn: TBitBtn;
    CloseBtn: TBitBtn;
    SumQuery: TADOQuery;
    SumDS: TDataSource;
    CopyBtn: TBitBtn;
    RegionPN: TPanel;
    RegionSB: TCheckBox;
    RegionCB: TComboBox;
    MainModePanel: TPanel;
    Panel3: TPanel;
    Panel7: TPanel;
    Panel13: TPanel;
    MainTypeCheck: TCheckBox;
    MainTypeCB: TCheckListBox;
    Label17: TLabel;
    MainModeCB: TComboBox;
    Label1: TLabel;
    DateCB: TComboBox;
    FirstDateED: TEdit;
    GetFirstDateBtn: TSpeedButton;
    LastDateED: TEdit;
    GetLastDateBtn: TSpeedButton;
    ProgressPn: TPanel;
    Label8: TLabel;
    ProgresCaption: TLabel;
    ProgressBar: TProgressBar;
    ProgresDate: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Panel2: TPanel;
    GroupCheck: TCheckBox;
    GroupCB: TCheckListBox;
    Panel14: TPanel;
    ReCheckBtn: TSpeedButton;
    UnCheckAllBtn: TSpeedButton;
    CheckAllBtn: TSpeedButton;
    ValLB: TLabel;
    ValCB: TComboBox;
    Label11: TLabel;
    FactoryCB: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ClearClientBtnClick(Sender: TObject);
    procedure UpdateFilter(Sender: TObject);
    function  UpdateGroupCB(mt : integer; update : boolean): integer;
    function  UpdateModelCB(mt : integer; update : boolean): integer;
    function  UpdateSenterCB: integer;
    function  UpdateCodeCB(mt : integer; update : boolean): integer;
    procedure FormShow(Sender: TObject);
    function  GetRecordFilterString(dt1,dt2 : TDate): string;
    function  GetReportFilterString(dt1,dt2 : TDate): string;
    function  GetCenterFilterString: string;
    procedure ChangeFilter(Sender: TObject);
    procedure PrepareBtnClick(Sender: TObject);
    procedure GetFirstDateBtnClick(Sender: TObject);
    procedure SetBtnClick(Sender: TObject);
    procedure CloseBtnClick(Sender: TObject);
    procedure CopyBtnClick(Sender: TObject);
    procedure ReportGetValue(const VarName: string; var Value: Variant);
    function  GetReportFilterCaption : string;
    function  GetRecordFilterCaption : string;
    procedure RegionCBChange(Sender: TObject);
    procedure RegionSBClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function  GetCheckedFromList(list : TCheckListBox): string;
    function  GetWorkCodeCheckList: string;
    procedure PrepareCellData(id : integer; dt1,dt2 : TDate);
    procedure FormResize(Sender: TObject);
    procedure CheckAllBtnClick(Sender: TObject);
    function  GetProdSQL(dt1,dt2 : TDateTime;id : integer):string;
  private
    { Private declarations }
    SumSql         : string;   //для сохранения запросов из компонентов
    FirstStart     : boolean;  //флаг первого запуска
    FirstDate      : TDate;
    LastDate       : TDate;
    MainGrid       : TMyGrid;
  public
    { Public declarations }
  end;


implementation

{$R *.dfm}

uses DataUnit, ClipBrd, {DateWin,} DateUtils, MonthWin{, PrintMod};




//--------------------- формирование и выполнение запросов ---------------------

function  TRepStatForMonthsForm.GetCenterFilterString : string;
var
  pid     : ^integer;
begin
  //отбор по региону
  if RegionSB.Checked then begin
    pid:=pointer(RegionCB.Items.Objects[RegionCB.ItemIndex]);
    result:='(REGIONID='+IntToStr(pid^)+')';
  end else result:='';
end;

function  TRepStatForMonthsForm.GetReportFilterString(dt1,dt2 : TDate) : string;
var
  flt,str : string;
  i       : integer;
  pid     : ^integer;
begin
  flt:='';
  //по дате, если дата отностится к отчету (отчетный период или дата отчета)
  if DateCB.ItemIndex<2 then begin
    case DateCB.ItemIndex of
      0 : str:='REPDATE';
      1 : str:='DOCDATE';
    end;
    flt:='('+str+' BETWEEN '+QuotedStr(FormatDateTime('yyyymmdd',StartOfTheDay(dt1))+' 00:00:00')+
      ' AND '+QuotedStr(FormatDateTime('yyyymmdd',EndOfTheDay(dt2))+' 23:59:59')+')';
  end;
  //по сервисным центрам
  str:='';
  if (SenterCheck.Checked) then begin
    str:='';
    for i := 0 to SenterCB.Items.Count-1 do
      if SenterCB.Checked[i] then begin
        pid:=pointer(SenterCB.Items.Objects[i]);
        if length(str)=0 then str:='(SENTERID='+inttostr(pid^)+')' else str:=str+' OR (SENTERID='+inttostr(pid^)+')';
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND ('+str+')';
  end;
  result:=flt;
end;

function  TRepStatForMonthsForm.GetRecordFilterString(dt1,dt2 : TDate) : string;
var
  flt,str : string;
  i       : integer;
  pid     : ^integer;
  pstr    : ^string;
begin
  flt:='';
  //по дате, если дата отностится к записи (дата производства, дата продажи, дата начала ремонта)
  if DateCB.ItemIndex>1 then begin
    case DateCB.ItemIndex of
      2 : str:='MAINDATE';
      3 : str:='BUYDATE';
      4 : str:='STARTDATE';
      5 : str:='ACCEPTDATE';
    end;
    flt:='('+str+' BETWEEN '+QuotedStr(FormatDateTime('yyyymmdd',StartOfTheDay(dt1)))
      +' AND '+QuotedStr(FormatDateTime('yyyymmdd',EndOfTheDay(dt2)))+')';
  end;
  //фабрика
  
  str:='';
  case FactoryCB.ItemIndex of
    1: str:='(FACTORYID=1)';
    2: str:='(FACTORYID=2)';
  end;
  if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  //вид ремонта
  if WorkTypeCheck.Checked then begin
    str:='(WORKTYPE='+QuotedStr(WorkTypeCB.Text)+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //срок службы
  if (WorkForClientCheck.Checked)and(WorkTypeCB.Visible)
    and(WorkTypeCB.ItemIndex=1)and(StrToIntDef(WorkForClientED.Text,0)>0 )then begin
      str:='(DATEDIFF(dd,BUYDATE,STARTDATE)<='+inttostr(strtointdef(WorkForClientEd.Text,0))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по данным клиента
  if (ClientDataCheck.Checked)and(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1) then begin
    if Length(ClientED.Text)>0 then begin
      str:='(CLIENT LIKE '+QuotedStr('%'+ClientED.Text+'%')+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if Length(AddrED.Text)>0 then begin
      str:='(CLIENTADDR LIKE '+QuotedStr('%'+AddrED.Text+'%')+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if Length(TelED.Text)>0 then begin
      str:='(CLIENTTEL LIKE '+QuotedStr('%'+TelED.Text+'%')+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
  end;
  //по типу продукции
 { str:='';
  if (self.MainModeCB.ItemIndex=3)and(MainTypeCheck.Checked) then begin
    str:='';
    for i := 0 to MainTypeCB.Items.Count-1 do
      if MainTypeCB.Checked[i] then begin
        pid:=pointer(MainTypeCB.Items.Objects[i]);
        if length(str)=0 then str:='(MAINTYPEID='+inttostr(pid^)+')' else str:=str+' OR (MAINTYPEID='+inttostr(pid^)+')';
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND ('+str+')';
  end;
  //по группам
  str:='';
  if (self.MainModeCB.ItemIndex<>2)and(GroupCB.Visible)and(GroupCheck.Checked)and(GroupCheck.Enabled) then begin
    str:='';
    for i := 0 to GroupCB.Items.Count-1 do
      if GroupCB.Checked[i] then begin
        pid:=pointer(GroupCB.Items.Objects[i]);
        if length(str)=0 then str:='(GROUPID='+inttostr(pid^)+')' else str:=str+' OR (GROUPID='+inttostr(pid^)+')';
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND ('+str+')';
  end;
  //по модели
  str:='';
  if (self.MainModeCB.ItemIndex<>2)and(ModelCB.Visible)and(ModelCheck.Checked)and(ModelCheck.Enabled) then begin
    str:='';
    for i := 0 to ModelCB.Items.Count-1 do
      if ModelCB.Checked[i] then begin
        pid:=pointer(ModelCB.Items.Objects[i]);
        if length(str)=0 then str:='(MODELID='+inttostr(pid^)+')' else str:=str+' OR (MODELID='+inttostr(pid^)+')';
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND ('+str+')';
  end; }
  //по серийному номеру
  if (SNCheck.Checked)and(Length(SNED.Text)>0) then begin
    str:='(SN LIKE '+QuotedStr('%'+SNED.Text+'%')+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по сроку ремонта
  if (WorkTimeCheck.Checked)and(StrToIntDef(WorkTimeED.Text,0)>0) then begin
    str:='((ENDDATE-STARTDATE)'+WorkTimeCB.Text+inttostr(strtointdef(WorkTimeED.Text,0))+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по стоимости
  if (CostCheck.Checked)then begin
    if(strtointdef(PartCostMin.Text,0)>0) then begin
      str:='(PARTCOST>'+ inttostr(strtointdef(PartCostMin.Text,0))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if(strtointdef(PartCostMax.Text,0)>0) then begin
      str:='(PARTCOST<='+inttostr(strtoint(PartCostMax.Text))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if (strtointdef(MovCostMin.Text,0)>0) then begin
      str:='(MOVPRICE>'+ inttostr(strtointdef(MovCostMin.Text,0))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if (strtointdef(MovCostMax.Text,0)>0) then begin
      str:='(MOVPRICE<='+inttostr(strtoint(MovCostMax.Text))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if(strtointdef(WorkCostMin.Text,0)>0) then begin
      str:='(WORKPRICE>'+ inttostr(strtointdef(WorkCostMin.Text,0))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
    if(strtointdef(WorkCostMax.Text,0)>0) then begin
      str:='(WORKPRICE<='+inttostr(strtoint(WorkCostMax.Text))+')';
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
    end;
  end;
  //по описанию детали
  if (PartCheck.Checked)and(Length(PartED.Text)>0) then begin
    str:='(PARTS LIKE '+QuotedStr('%'+PartED.Text+'%')+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по описанию проблемы
  if (ProblemNoteCheck.Checked)and(Length(ProblemNoteED.Text)>0) then begin
    str:='(PROBLEMNOTE LIKE '+QuotedStr('%'+ProblemNoteED.Text+'%')+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по описанию работ
  if (WorkNoteCheck.Checked)and(Length(WorkNoteED.Text)>0) then begin
    str:='(WORKNOTE LIKE '+QuotedStr('%'+WorkNoteED.Text+'%')+')';
    if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND '+str;
  end;
  //по коду неисправности
  str:='';
  if (self.MainModeCB.ItemIndex<>3)and(WorkCodeCB.Visible)and(WorkCodeCheck.Checked)and(WorkCodeCheck.Enabled) then begin
    str:='';
    for i := 0 to WorkCodeCB.Items.Count-1 do
      if WorkCodeCB.Checked[i] then begin
        pstr:=pointer(WorkCodeCB.Items.Objects[i]);
        if length(str)=0 then str:='(WORKCODE='+QuotedStr(pstr^)+')' else str:=str+' OR (WORKCODE='+QuotedStr(pstr^)+')';
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+chr(13)+' AND ('+str+')';
  end;
  result:=flt;
end;

procedure TRepStatForMonthsForm.PrepareCellData(id : integer; dt1,dt2 : TDate);
var
  sql, filter, recflt, repflt,sntflt : string;
begin
  recflt:=self.GetRecordFilterString(dt1,dt2);
  repflt:=self.GetReportFilterString(dt1,dt2);
  sntflt:=self.GetCenterFilterString;
  filter:='';
  case self.MainModeCB.ItemIndex of
    0 : filter:=' AND (MAINTYPEID='+inttostr(id)+')';
    1 : filter:=' AND (GROUPID='+inttostr(id)+')';
    2 : filter:=' AND (MODELID='+inttostr(id)+')';
  end;
  if Length(recflt)>0 then filter:=filter+chr(13)+'AND ('+recflt+')';
  if (Length(repflt)>0)or(Length(sntflt)>0) then begin
    filter:=filter+chr(13)+('AND [REPORTID] IN (SELECT ID FROM [SERVREPORT] WHERE '+repflt);
    if Length(sntflt)=0 then filter:=filter+')' else begin
      if Length(repflt)>0 then filter:=filter+'AND ';
      filter:=filter+'[SENTERID] IN (SELECT ID FROM [SERVCENTRES] WHERE '+SNTFLT+'))';
    end;
  end;
  SumQuery.Close;
  SumQuery.SQL.Clear;
  sql:=stringreplace(SumSQL,'/FILTER/',filter,[rfReplaceAll]);
  if (self.ValCB.ItemIndex>0) then
    sql:=stringreplace(SQL,'/PRODFILTER/',self.GetProdSQL(dt1,dt2,id),[rfReplaceAll])
      else sql:=stringreplace(SQL,'/PRODFILTER/','',[rfReplaceAll]);
  SumQuery.SQL.Add(sql);

  //SHOWMESSAGE(sql);
  //Clipboard.AsText:=sql;

  SumQuery.Open;
end;

function  TRepStatForMonthsForm.GetWorkCodeCheckList: string;
var
  i    : integer;
  pstr : ^string;
  str  : string;
begin
    str:='';
    for i := 0 to WorkCodeCB.Items.Count-1 do
      if WorkCodeCB.Checked[i] then begin
        pstr:=pointer(WorkCodeCB.Items.Objects[i]);
        if length(str)=0 then str:='(T1.CODE='+QuotedStr(pstr^)+')' else str:=str+' OR (T1.CODE='+QuotedStr(pstr^)+')';
      end;
    result:=str;
end;

function  TRepStatForMonthsForm.GetCheckedFromList(list : TCheckListBox):string;
var
  i    : integer;
  pint : ^integer;
begin
  result:='';
  for I := 0 to List.Count - 1 do
    if (List.Checked[i])and(List.Items.Objects[i]<>nil) then begin
        pint:=pointer(List.Items.Objects[i]);
        result:=result+inttostr(pint^)+',';
      end;
  if length(result)>0 then result:='('+copy(result,1,length(result)-1)+')';
end;

procedure TRepStatForMonthsForm.PrepareBtnClick(Sender: TObject);
var
  i,c,r,prod : integer;
  RowQuery   : TADOQuery;
  mainfilter,
  str        : string;
  sql        : TStringList;
  pid        : ^integer;
begin
  //формирование текста запроса для заполнения заголовка строк
  sql      := TStringList.Create;
  RowQuery := TADOQuery.Create(self);
  RowQuery.Connection:=DataMod.Connection;
  case self.MainModeCB.ItemIndex of
    0 : begin
          sql.Add('SELECT T1.ID AS ID, T1.DESCR FROM [MAINTYPES] AS T1');
          if self.MainTypeCheck.Checked then begin
            mainfilter:=GetCheckedFromList(self.MainTypeCB);
            if length(mainfilter)>0 then sql.Add('WHERE T1.[ID] IN '+mainfilter);
          end;
          sql.Add('ORDER BY T1.DESCR');
          //проверка выбора хотябы одного элемента в списке выбора
          i:=0;
          while(i<MainTypeCB.Count)and(not MainTypeCB.Checked[i])do inc(i);
          if(i>=MainTypeCB.Count)or(not MainTypeCheck.Checked)then
            if MessageDlg('Ни один тип продукции не выбран!'+chr(13)+
              'Будут выбраны все типы.',mtWarning,[mbOk,mbCancel],0)=mrCancel then Abort;
        end;
    1 : begin
          sql.Add('SELECT T1.ID AS ID, T1.MAINTYPE AS MT, CONCAT(T2.DESCR,'+QuotedStr(' ')+',T1.DESCR) AS DESCR ,MAINTYPE');
          sql.Add('FROM [PRODUCTIONGROUPS] AS T1');
          sql.Add('LEFT JOIN [MAINTYPES] AS T2 ON T2.[ID]=T1.[MAINTYPE]');
          sql.Add('WHERE (LEN(T1.DESCR)>0)');
          if self.MainTypeCheck.Checked then begin
            mainfilter:=GetCheckedFromList(self.MainTypeCB);
            if length(mainfilter)>0 then mainfilter:=' AND (T2.[ID] IN '+mainfilter+')';
          end;
          if (GroupCheck.Checked)and(GroupCheck.Enabled) then begin
            str:=GetCheckedFromList(self.GroupCB);
            if length(str)>0 then mainfilter:=mainfilter+' AND (T1.[ID] IN '+str+')';
          end;
          if length(mainfilter)>0 then sql.Add(mainfilter);
          sql.Add('ORDER BY DESCR');
          //проверка выбора хотябы одного элемента в списке выбора
          i:=0;
          while(i<GroupCB.Count)and(not GroupCB.Checked[i])do inc(i);
          if(i>=GroupCB.Count)or(not GroupCheck.Checked)then
            if MessageDlg('Ни одна группа прдукции не выбрана!'+chr(13)+
              'Будут выбраны все группы.',mtWarning,[mbOk,mbCancel],0)=mrCancel then Abort;
        end;
    2 : begin
          sql.Add('SELECT T1.ID AS ID, T1.MAINTYPE AS MT, CONCAT(T2.DESCR,'+QuotedStr(' ')+',T1.DESCR) AS DESCR ,MAINTYPE');
          sql.Add('FROM [PRODUCTIONMODELS] AS T1');
          sql.Add('LEFT JOIN [MAINTYPES] AS T2 ON T2.[ID]=T1.[MAINTYPE]');
          sql.Add('WHERE (LEN(T1.DESCR)>0)');
          if self.MainTypeCheck.Checked then begin
            mainfilter:=GetCheckedFromList(self.MainTypeCB);
            if length(mainfilter)>0 then mainfilter:=' AND (T2.[ID] IN '+mainfilter+')';
          end;
          if (ModelCheck.Checked)and(ModelCheck.Enabled) then begin
            str:=GetCheckedFromList(self.ModelCB);
            if length(str)>0 then mainfilter:=mainfilter+' AND (T1.[ID] IN '+str+')';
          end;
          if length(mainfilter)>0 then sql.Add(mainfilter);
          sql.Add('ORDER BY DESCR');
          //проверка выбора хотябы одного элемента в списке выбора
          i:=0;
          while(i<ModelCB.Count)and(not ModelCB.Checked[i])do inc(i);
          if(i>=ModelCB.Count)or(not ModelCheck.Checked)then
            if MessageDlg('Ни одина модель не выбрана!'+chr(13)+
              'Будут выбраны все модели.',mtWarning,[mbOk,mbCancel],0)=mrCancel then Abort;
        end;
  end;
  //Заполнение закголовков строк
  RowQuery.SQL.Clear;
  RowQuery.SQL.AddStrings(sql);
  RowQuery.Open;
  MainGrid.RepaintHeader(FirstDate,LastDate,tmMonth);
  MainGrid.RowCount:=MainGrid.FixedRows+RowQuery.RecordCount;
  r:=MainGrid.FixedRows;
  while not RowQuery.Eof do begin
    MainGrid.Cells[1,r]:=RowQuery.FieldByName('DESCR').AsString;
    if MainGrid.Objects[0,r]<>nil then FreeMem(pointer(MainGrid.Objects[0,r]));
    new(pid);
    pid^:=RowQuery.FieldByName('ID').AsInteger;
    MainGrid.Objects[0,r]:=TObject(pid);
    inc(r);
    RowQuery.Next;
  end;
  MainGrid.Visible:=true;
  sql.Free;
  RowQuery.Free;
  if FirstStart then FirstStart:=false;
  //
  ProgressPn.Left:=round((self.Width-ProgressPn.Width)/2);
  ProgressPn.Top:=round((self.Height-ProgressPn.Height*2)/2);
  ProgressPn.BringToFront;
  ProgressPn.Visible:=true;
  ProgressBar.Max:=(MainGrid.RowCount-MainGrid.FixedRows)*(MainGrid.ColCount-MainGrid.FixedCols);
  ProgressBar.Position:=0;
  ProgresCaption.Caption:='';
  ProgresDate.Caption:='';
  ProgressPn.Repaint;
  //вывод значений в ячейки
  for r := MainGrid.FixedRows to MainGrid.RowCount - 1 do
    if MainGrid.Objects[0,r]<>nil then begin
      ProgresCaption.Caption:='Строка: '+MainGrid.Cells[1,r];
      //расчет общих итогов
      ProgresDate.Caption:='Расчет общих итогов.';
      pid:=pointer(MainGrid.Objects[0,r]);
      self.PrepareCellData(pid^,MainGrid.TimePeriod.Period[0].first,
        MainGrid.TimePeriod.Period[MainGrid.TimePeriod.PeriodsCount-1].last);
      str:=SumQuery.FieldByName('CNT').AsString;
      if (ValCB.Visible)and(ValCB.ItemIndex>0) then begin
        prod:=SumQuery.FieldByName('PROD').AsInteger;
        case ValCB.ItemIndex of
          1 : str:=str+'/'+SumQuery.FieldByName('PROD').AsString;
          2 : if prod=0 then str:=str+'/0'
                else str:=str+'/'+FormatFloat('##0.00',SumQuery.FieldByName('CNT').AsInteger/prod*100)+'%';
          3 : if prod=0 then str:='0'
                else str:=inttostr(prod);
        end;
      end;
      MainGrid.Cells[MainGrid.FixedCols,r]:=str;
      //расчет итогов по периоду
      i:=0;
      for c :=MainGrid.FixedCols+1  to MainGrid.ColCount - 1 do begin
        ProgresDate.Caption:='Расчет за период: '+FormatDateTime('dd.mm.yy',MainGrid.TimePeriod.Period[i].first)+
          ' - '+FormatDateTime('dd.mm.yy',MainGrid.TimePeriod.Period[i].last);
        ProgressBar.Position:=ProgressBar.Position+1;
        ProgressPn.Repaint;
        pid:=pointer(MainGrid.Objects[0,r]);
        self.PrepareCellData(pid^,MainGrid.TimePeriod.Period[i].first,
          MainGrid.TimePeriod.Period[i].last);
        str:=SumQuery.FieldByName('CNT').AsString;
        if (ValCB.Visible)and(ValCB.ItemIndex>0) then begin
          prod:=SumQuery.FieldByName('PROD').AsInteger;
          case ValCB.ItemIndex of
            1 : str:=str+'/'+SumQuery.FieldByName('PROD').AsString;
            2 : if prod=0 then str:=str+'/0'
                else str:=str+'/'+FormatFloat('##0.00',SumQuery.FieldByName('CNT').AsInteger/prod*100)+'%';
            3 : if prod=0 then str:='0'
                else str:=inttostr(prod);
          end;
        end;
        MainGrid.Cells[c,r]:=str;
        inc(i);
      end;
    end;
  ProgressPn.Visible:=false;
  MainGrid.Font.Color:=clBlack;
end;

function  TRepStatForMonthsForm.GetProdSQL(dt1,dt2 : TDateTime;id : integer):string;
var
  str       : string;
  modtbl    : TADOQuery;
  txt       : TStringList;
begin
  Txt:=TStringList.Create;
  txt.Add('SELECT @PROD=COUNT(DISTINCT T1.[CODE]) FROM [NOVATEKPRODCODES] AS T1');
  txt.Add('WHERE (T1.[DATETIME] BETWEEN '+QuotedStr(FormatDateTime('yyyymmdd',StartOfTheDay(dt1))+' 00:00:00')+
      ' AND '+QuotedStr(FormatDateTime('yyyymmdd',EndOfTheDay(dt2))+' 23:59:59')+')');
  //
  case self.MainModeCB.ItemIndex of
    0 : txt.Add('AND (MAINTYPE='+inttostr(id)+')');
    1 : begin
          modtbl:=TADOQuery.Create(self);
          modtbl.Connection:=DataMod.Connection;
          case FactoryCB.ItemIndex of
            1: modtbl.SQL.Add('SELECT IDFORCODES AS ID FROM PRODUCTIONMODELS');
            2: modtbl.SQL.Add('SELECT IDFORCODESEAST AS ID FROM PRODUCTIONMODELS');
          end;
          modtbl.SQL.Add('WHERE PRODUCTGROUP='+IntToStr(ID));
          modtbl.Open;
          str:='';
          while not modtbl.Eof do begin
            str:=str+'(CODE LIKE '+QuotedStr(modtbl.FieldByName('ID').AsString+'%')+')OR';
            modtbl.Next;
          end;
          if length(str)>0 then begin
            str:=copy(str,1,length(str)-2);
            txt.Add('AND ('+str+')');
          end;
          modtbl.Free;
        end;
    2 : begin
          modtbl:=TADOQuery.Create(self);
          modtbl.Connection:=DataMod.Connection;
          case FactoryCB.ItemIndex of
            1: modtbl.SQL.Add('SELECT IDFORCODES AS ID FROM PRODUCTIONMODELS');
            2: modtbl.SQL.Add('SELECT IDFORCODESEAST AS ID FROM PRODUCTIONMODELS');
          end;
          modtbl.SQL.Add('WHERE ID='+IntToStr(ID));
          modtbl.Open;
          if not modtbl.IsEmpty then txt.Add('AND (CODE LIKE '+QuotedStr(modtbl.FieldByName('ID').AsString+'%')+')');
          modtbl.Free;
        end;
  end;
  //фабрика
  case FactoryCB.ItemIndex of
    1 : txt.Add('AND ((SHOPSID=1)OR(SHOPSID=2)OR(SHOPSID=3))');
    2 : txt.Add('AND SHOPSID=4)');
  end;
  result:=txt.Text;
  //showmessage(result);
  txt.Free;
end;

//-------------------

function  TRepStatForMonthsForm.GetReportFilterCaption : string;
var
  flt,str : string;
  i       : integer;
begin
  flt:='';
  //по дате, если дата отностится к отчету (отчетный период или дата отчета)
  if DateCB.ItemIndex<2 then begin
    case DateCB.ItemIndex of
      0 : str:='отчетный период';
      1 : str:='дата отчета';
    end;
    flt:=str+' с '+QuotedStr(FormatDateTime('mmmm yyyy',FirstDate))+' по '+QuotedStr(FormatDateTime('mmmm yyyy',LastDate));
  end;
  //по региону
  if RegionSB.Checked then
    if length(flt)=0 then flt:='регион СЦ - '+RegionCB.Text else flt:=flt+', регион СЦ - '+RegionCB.Text;
  //по сервисным центрам
  str:='';
  if (SenterCheck.Checked) then begin
    str:='';
    for i := 0 to SenterCB.Items.Count-1 do
      if SenterCB.Checked[i] then
        if length(str)=0 then str:='сервисные центры: '+SenterCB.Items[i] else str:=str+', '+SenterCB.Items[i];
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  result:=flt;
end;

function  TRepStatForMonthsForm.GetRecordFilterCaption : string;
var
  flt,str : string;
  i       : integer;
  pid     : ^integer;
  pstr    : ^string;
begin
  flt:='';
  //по дате, если дата отностится к записи (дата производства, дата продажи, дата начала ремонта)
  if DateCB.ItemIndex>1 then begin
    case DateCB.ItemIndex of
      2 : str:='дата производства';
      3 : str:='дата продажи';
      4 : str:='дата начала ремонта';
    end;
    flt:=str+' с '+QuotedStr(FormatDateTime('mmmm yyyy',FirstDate))+' по '+QuotedStr(FormatDateTime('mmmm yyyy',LastDate));
  end;
  //вид ремонта
  if WorkTypeCheck.Checked then begin
    str:='тип ремонта '+QuotedStr(WorkTypeCB.Text);
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //срок службы
  if (WorkForClientCheck.Checked)and(WorkTypeCB.Visible)
    and(WorkTypeCB.ItemIndex=1)and(StrToIntDef(WorkForClientED.Text,0)>0 )then begin
      str:='срок службы до ремонта '+WorkForClientEd.Text+' дней';
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по данным клиента
  if (ClientDataCheck.Checked)and(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1) then begin
    if Length(ClientED.Text)>0 then begin
      str:='по имени клиента '+QuotedStr('%'+ClientED.Text+'%');
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if Length(AddrED.Text)>0 then begin
      str:='по адресу клиента '+QuotedStr('%'+AddrED.Text+'%');
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if Length(TelED.Text)>0 then begin
      str:='по телефону клиента '+QuotedStr('%'+TelED.Text+'%');
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
  end;
  //по типу продукции
  str:='';
  if (MainTypeCheck.Checked) then begin
    str:='';
    for i := 0 to MainTypeCB.Items.Count-1 do
      if MainTypeCB.Checked[i] then
        if length(str)=0 then str:='по типу прдукции '+QuotedStr(MainTypeCB.Items[i])
          else str:=str+', '+QuotedStr(MainTypeCB.Items[i]);
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по модели
  str:='';
  if (ModelCB.Visible)and(ModelCheck.Checked)and(ModelCheck.Enabled) then begin
    str:='';
    for i := 0 to ModelCB.Items.Count-1 do
      if ModelCB.Checked[i] then begin
        pid:=pointer(ModelCB.Items.Objects[i]);
        if length(str)=0 then str:='по модели '+QuotedStr(ModelCB.Items[i]) else str:=str+', '+QuotedStr(ModelCB.Items[i]);
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по серийному номеру
  if (SNCheck.Checked)and(Length(SNED.Text)>0) then begin
    str:='по серийному номеру '+QuotedStr('%'+SNED.Text+'%');
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по сроку ремонта
  if (WorkTimeCheck.Checked)and(StrToIntDef(WorkTimeED.Text,0)>0) then begin
    str:='по сроку ремонта '+WorkTimeCB.Text+WorkTimeED.Text+'дней ';
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по стоимости
  if (CostCheck.Checked)then begin
    if(strtointdef(PartCostMin.Text,0)>0) then begin
      str:='цена детали>'+PartCostMin.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if(strtointdef(PartCostMax.Text,0)>0) then begin
      str:='цена детали<='+PartCostMax.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if (strtointdef(MovCostMin.Text,0)>0) then begin
      str:='за выезд>'+MovCostMin.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if (strtointdef(MovCostMax.Text,0)>0) then begin
      str:='за выезд<='+MovCostMax.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if(strtointdef(WorkCostMin.Text,0)>0) then begin
      str:='за ремонт>'+WorkCostMin.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
    if(strtointdef(WorkCostMax.Text,0)>0) then begin
      str:='за ремонт<='+WorkCostMax.Text;
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
    end;
  end;
  //по описанию детали
  if (PartCheck.Checked)and(Length(PartED.Text)>0) then begin
    str:='по описанию детали '+QuotedStr('%'+PartED.Text+'%');
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по описанию проблемы
  if (ProblemNoteCheck.Checked)and(Length(ProblemNoteED.Text)>0) then begin
    str:='(по описанию проблемы '+QuotedStr('%'+ProblemNoteED.Text+'%');
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по описанию работ
  if (WorkNoteCheck.Checked)and(Length(WorkNoteED.Text)>0) then begin
    str:='по описанию работ '+QuotedStr('%'+WorkNoteED.Text+'%');
    if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  //по коду неисправности
  str:='';
  if (WorkCodeCB.Visible)and(WorkCodeCheck.Checked)and(WorkCodeCheck.Enabled) then begin
    str:='';
    for i := 0 to WorkCodeCB.Items.Count-1 do
      if WorkCodeCB.Checked[i] then begin
        pstr:=pointer(WorkCodeCB.Items.Objects[i]);
        if length(str)=0 then str:='по коду неисправности '+QuotedStr(pstr^) else str:=str+', '+QuotedStr(pstr^);
      end;
    if length(str)>0 then
      if length(flt)=0 then flt:=str else flt:=flt+', '+str;
  end;
  result:=flt;
end;

//-------------------события контролов ----------------------------------------

procedure TRepStatForMonthsForm.ClearClientBtnClick(Sender: TObject);
begin
  if TSpeedButton(sender).Name='ClearClientBtn' then ClientED.Text:='';
  if TSpeedButton(sender).Name='ClearAddrBtn' then AddrED.Text:='';
  if TSpeedButton(sender).Name='ClearTelBtn' then TelED.Text:='';
  if TSpeedButton(sender).Name='ClearSNEDBTN' then SNED.Text:='';
  if TSpeedButton(sender).Name='ClearWorkTimeBtn' then WorkTimeED.Text:='';
  if TSpeedButton(sender).Name='ClearPartBtn' then PartED.Text:='';
  if TSpeedButton(sender).Name='ClearProblemNoteBtn' then ProblemNoteED.Text:='';
  if TSpeedButton(sender).Name='ClearWorkNoteBtn' then WorkNoteED.Text:='';
  self.ChangeFilter(self);
end;

procedure TRepStatForMonthsForm.CloseBtnClick(Sender: TObject);
begin
  self.Close;
end;

procedure TRepStatForMonthsForm.CopyBtnClick(Sender: TObject);
var
  c,r:integer;
begin
  if (MainGrid.Selection.Top<0)or(MainGrid.Selection.Left<0) then Abort;
  Clipboard.Open;
  Clipboard.Clear;
  for r := MainGrid.Selection.Top to MainGrid.Selection.Bottom do begin
    for c := MainGrid.Selection.Left  to MainGrid.Selection.Right do begin
        Clipboard.AsText:=Clipboard.AsText+MainGrid.Cells[c,r];
        if c<MainGrid.Selection.Right then Clipboard.AsText:=Clipboard.AsText+#9;
      end;
    Clipboard.AsText:=Clipboard.AsText+#13#10;
  end;
  Clipboard.Close;
end;

procedure TRepStatForMonthsForm.ReportGetValue(const VarName: string; var Value: Variant);
var
  str : string;
begin
  if comparetext(varname,'FILTER')=0 then begin
    //формирование заголовка с данными о фильтр
    value:='';
    str:=self.GetReportFilterCaption;
    if length(str)>0 then value:='Фильтр по отчептам:'+chr(13)+str+chr(13);
    str:=self.GetRecordFilterCaption;
    if length(str)>0 then value:=value+'Фильтр по ремонтам:'+chr(13)+str+chr(13);
  end;
end;

procedure TRepStatForMonthsForm.SetBtnClick(Sender: TObject);
begin
  MainPN.Visible:=SetBtn.Down;
  SB.VertScrollBar.Position:=0;
end;

//-------------------- контролы настроек фильтра --------------------------------

procedure TRepStatForMonthsForm.GetFirstDateBtnClick(Sender: TObject);
var
  month,year : word;
  txt : string;
begin
  if (sender as TControl).Name='GetFirstDateBtn' then begin
    Year:=YearOf(FirstDate);
    Month:=MonthOf(FirstDate);
    txt:=GetMonth(month,year);
    if length(txt)>0 then begin
      FirstDateEd.Text:=txt;
      FirstDate:=EncodeDate(year,month,1);
    end;
  end;
  if (sender as TControl).Name='GetLastDateBtn' then begin
    Year:=YearOf(LastDate);
    Month:=MonthOf(LastDate);
    txt:=GetMonth(month,year);
    if length(txt)>0 then begin
      LastDateEd.Text:=txt;
      LastDate:=EndOfTheMonth(EncodeDate(year,month,1));
    end;
  end;
  self.ChangeFilter(self);
end;

procedure TRepStatForMonthsForm.ChangeFilter(Sender: TObject);
begin
  MainGrid.Font.Color:=clGray;
  MainGrid.Repaint;
end;

procedure TRepStatForMonthsForm.UpdateFilter(Sender: TObject);
var
  i,j : integer;
  pid : ^integer;
  upd : boolean;
begin
  ValLB.Visible:=(FactoryCB.ItemIndex>0)AND(DateCB.ItemIndex=2);
  ValCB.Visible:=ValLB.Visible;
  RegionCB.Visible:=RegionSB.Checked;
  SenterCB.Visible:=SenterCheck.Checked;
  WorkTypeCB.Visible:=(WorkTypeCheck.Checked);
  WorkForClientCheck.Enabled:=(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1);
  WorkForClientEd.Visible:=(WorkForClientCheck.Checked)and(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1);
  ClientDataCheck.Enabled:=(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1);
  ClientDataPn.Visible:=(ClientDataCheck.Checked)and(WorkTypeCB.Visible)and(WorkTypeCB.ItemIndex=1);
  MainTypeCB.Visible:=(MainTypeCheck.Checked);
  //считаем кол-во выбранных типов
  upd:=((sender as TControl).Name='MainTypeCB')or((sender as TControl).Name='MainTypeCheck');
  j:=0;
  pid:=nil;
  for i := 0 to MainTypeCB.Items.Count-1 do
    if MainTypeCB.Checked[i] then begin
      inc(j);
      pid:=pointer(MainTypeCB.Items.Objects[i]);
    end;
  //если тип один заполняем список групп и считаем ко-во групп
  i:=0;
  if (j=1)and(pid<>nil) then i:=self.UpdateGroupCB(pid^,upd);
  GroupCheck.Enabled:=(i>0)and(j=1)and(MainTypeCheck.Checked)and(MainModeCB.ItemIndex>=1);
  GroupCB.Visible:=(GroupCheck.Checked)and(i>0)and(j=1)and(GroupCheck.Enabled)and(MainModeCB.ItemIndex>=1);
  //если тип один заполняем список моделей и считаем ко-во моделей
  i:=0;
  if (j=1)and(pid<>nil) then
    if pos('Group',(sender as TControl).Name)>0 then i:=self.UpdateModelCB(pid^,true)
      else i:=self.UpdateModelCB(pid^,upd);
  ModelCheck.Enabled:=(i>0)and(j=1)and(MainTypeCheck.Checked)and(MainModeCB.ItemIndex=2);
  ModelCB.Visible:=(ModelCheck.Checked)and(i>0)and(j=1)and(ModelCheck.Enabled)and(MainModeCB.ItemIndex=2);
  //если тип один заполняем список кодов и считаем ко-во кодов
  i:=0;
  if (j=1)and(pid<>nil) then i:=self.UpdateCodeCB(pid^,upd) else i:=self.UpdateCodeCB(0,true);
  WorkCodeCheck.Enabled:=(i>0);
  WorkCodeCB.Visible:=(WorkCodeCheck.Checked)and(i>0);
  //
  SNED.Visible:=SNCheck.Checked;
  ClearSNEDBtn.Visible:=SNCheck.Checked;
  WorkTimeED.Visible:=WorkTimeCheck.Checked;
  WorkTimeCB.Visible:=WorkTimeCheck.Checked;
  ClearWorkTimeBtn.Visible:=WorkTimeCheck.Checked;
  CostDataPn.Visible:=CostCheck.Checked;
  PartED.Visible:=PartCheck.Checked;
  ClearPartBtn.Visible:=PartCheck.Checked;
  ProblemNoteED.Visible:=ProblemNoteCheck.Checked;
  ClearProblemNoteBtn.Visible:=ProblemNoteCheck.Checked;
  WorkNoteED.Visible:=WorkNoteCheck.Checked;
  ClearWorkNoteBtn.Visible:=WorkNoteCheck.Checked;
  SB.VertScrollBar.Range:=BottomBevel.Top+BottomBevel.Height;
  self.ChangeFilter(self);
end;

procedure TRepStatForMonthsForm.RegionCBChange(Sender: TObject);
begin
  if self.UpdateSenterCB=0 then begin
    SenterCheck.Checked:=false;
    SenterCheck.Enabled:=false;
    SenterCB.Visible:=false;
  end else SenterCheck.Enabled:=true;
  self.ChangeFilter(sender);
end;

procedure TRepStatForMonthsForm.RegionSBClick(Sender: TObject);
begin
  self.UpdateFilter(sender);
  self.RegionCBChange(sender);
end;

function  TRepStatForMonthsForm.UpdateGroupCB(mt : integer; update : boolean):integer;
var
  i : integer;
  pid : ^integer;
  Table : TADOTable;
begin
  Table := TADOTable.Create(self);
  Table.Connection:=DataMod.Connection;
  Table.TableName:='PRODUCTIONGROUPS';
  Table.Open;
  if update then GroupCB.Clear;
  i:=0;
  while not Table.Eof do begin
    if (Table.FieldByName('MAINTYPE').AsInteger=mt)and
      (length(Table.FieldByName('DESCR').AsString)>0)then begin
        if update then begin
          GroupCB.Items.Add(Table.FieldByName('DESCR').AsString);
          new(pid);
          pid^:=Table.FieldByName('ID').AsInteger;
          GroupCB.Items.Objects[GroupCB.Items.Count-1]:=TObject(pid);
        end;
        inc(i);
      end;
    Table.Next;
  end;
  Table.Free;
  result:=i;
end;

function  TRepStatForMonthsForm.UpdateModelCB(mt : integer; update : boolean):integer;
var
  i : integer;
  pid : ^integer;
  ModelsTable : TADOQuery;
  str : string;
begin
  ModelsTable := TADOQuery.Create(self);
  ModelsTable.Connection:=DataMod.Connection;
  ModelsTable.SQL.Add('SELECT * FROM PRODUCTIONMODELS WHERE MAINTYPE='+INTTOSTR(MT));
  if (GroupCheck.Checked)and(GroupCheck.Enabled) then begin
    str:=self.GetCheckedFromList(self.GroupCB);
    if Length(str)>0 then ModelsTable.SQL.Add('AND PRODUCTGROUP IN '+str);
  end;
  ModelsTable.SQL.Add('ORDER BY PRODUCTGROUP, DESCR');
  ModelsTable.Open;
  if update then ModelCB.Clear;
  i:=0;
  while not ModelsTable.Eof do begin
    if (ModelsTable.FieldByName('MAINTYPE').AsInteger=mt)and
      (length(ModelsTable.FieldByName('DESCR').AsString)>0)then begin
        if update then begin
          ModelCB.Items.Add(ModelsTable.FieldByName('DESCR').AsString);
          new(pid);
          pid^:=ModelsTable.FieldByName('ID').AsInteger;
          ModelCB.Items.Objects[ModelCB.Items.Count-1]:=TObject(pid);
        end;
        inc(i);
      end;
    ModelsTable.Next;
  end;
  ModelsTable.Free;
  result:=i;
end;

function  TRepStatForMonthsForm.UpdateSenterCB:integer;
var
  i : integer;
  pid : ^integer;
  Query : TADOQuery;
begin
  Query:=TADOQuery.Create(self);
  Query.Connection:=DataMod.Connection;
  Query.SQL.Add('SELECT * FROM SERVCENTRES');
  if RegionSB.Checked then begin
    pid:=pointer(RegionCB.Items.Objects[RegionCB.ItemIndex]);
    Query.SQL.Add('WHERE REGIONID='+inttostr(pid^));
  end;
  Query.Open;
  SenterCB.Clear;
  i:=0;
  while not Query.Eof do begin
    SenterCB.Items.Add(Query.FieldByName('DESCR').AsString);
    new(pid);
    pid^:=Query.FieldByName('ID').AsInteger;
    SenterCB.Items.Objects[SenterCB.Items.Count-1]:=TObject(pid);
    inc(i);
    Query.Next;
  end;
  Query.Free;
  result:=i;
end;

function  TRepStatForMonthsForm.UpdateCodeCB(mt : integer; update : boolean):integer;
var
  i : integer;
  pid : ^string;
  Query : TADOQuery;
begin
  Query:=TADOQuery.Create(self);
  Query.Connection:=DataMod.Connection;
  if mt>0 then Query.SQL.Add('SELECT * FROM SERVCODES WHERE ((MAINTYPE='+INTTOSTR(MT)+')OR(MAINTYPE=0))AND(ISFOLDER=0)'+
    chr(13)+'ORDER BY MAINTYPE,CODE')
    else Query.SQL.Add('SELECT * FROM SERVCODES WHERE (MAINTYPE=0) AND(ISFOLDER=0)');
  Query.Open;
  if(update)then WorkCodeCB.Clear;
  i:=0;
  while not Query.Eof do begin
    if(update)then begin
      new(pid);
      pid^:=Query.FieldByName('CODE').AsString;
      WorkCodeCB.Items.Add(Query.FieldByName('CODE').AsString+' '+Query.FieldByName('DESCR').AsString);
      WorkCodeCB.Items.Objects[WorkCodeCB.Items.Count-1]:=TObject(pid);
    end;
    inc(i);
    Query.Next;
  end;
  Query.Free;
  result:=i;
end;

procedure TRepStatForMonthsForm.CheckAllBtnClick(Sender: TObject);
var
  box : TCheckListBox;
  i   : integer;
begin
  if (self.ActiveControl is TCheckListBox) then begin
    box:=(self.ActiveControl as TCheckListBox);
    for I := 0 to box.Count - 1 do
      case (sender as TSpeedButton).Tag of
        1 : box.Checked[i]:=true;
        2 : box.Checked[i]:=false;
        3 : box.Checked[i]:=not box.Checked[i];
      end;
  end;
end;

//---------------------- события формы -----------------------------------------

procedure TRepStatForMonthsForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  PostMessage(application.MainForm.Handle,WM_CLOSECHILD,self.Tag,0);
  Action:=caFree;
end;

procedure TRepStatForMonthsForm.FormCreate(Sender: TObject);
begin
  MainGrid:=TMyGrid.Create(self);
  MainGrid.Align:=alClient;
  MainGrid.ColWidths[1]:=200;
  MainGrid.Visible:=false;
  Panel1.InsertControl(MainGrid);
end;

procedure TRepStatForMonthsForm.FormResize(Sender: TObject);
begin
  SB.VertScrollBar.Position:=0;
end;

procedure TRepStatForMonthsForm.FormShow(Sender: TObject);
var
  pid   : ^integer;
  Table : TADOTable;
begin
  SumSQL:=string(SumQuery.SQL.Text);
  //Позиционирование контролов (для разных экранов)
  RegionCB.Top:=round((RegionPN.Height-RegionCB.Height)/2)-3;
  WorkTypeCB.Top:=round((Panel4.Height-WorkTypeCB.Height)/2)-3;
  WorkForClientED.Top:=round((WorkForClientPn.Height-WorkForClientED.Height)/2)-3;
  SNEd.Top:=round((Panel5.Height-SNed.Height)/2)-3;
  ClearSNEDBTN.Top:=SNEd.Top;
  WorkTimeCB.Top:=round((Panel8.Height-WorkTimeCB.Height)/2)-3;
  WorkTimeED.Top:=WorkTimeCB.Top;
  ClearWorkTimeBtn.Top:=WorkTimeCB.Top;
  PartED.Top:=round((Panel9.Height-PartEd.Height)/2)-3;
  ClearPartBtn.Top:=PartED.Top;
  ProblemNoteED.Top:=round((Panel10.Height-ProblemNoteEd.Height)/2)-3;
  ClearProblemNoteBtn.Top:=ProblemNoteED.Top;
  WorkNoteED.Top:=round((Panel11.Height-WorkNoteEd.Height)/2)-3;
  ClearWorkNoteBtn.Top:=WorkNoteED.Top;
  //заполняем ComboBox выбора региона
  RegionCB.Clear;
  RegionCB.Sorted:=false;
  Table := TADOTable.Create(self);
  Table.Connection:=DataMod.Connection;
  Table.TableName:='SERVREGION';
  Table.Open;
  while not Table.Eof do begin
    RegionCB.Items.Add(Table.FieldByName('DESCR').AsString);
    new(pid);
    pid^:=Table.FieldByName('ID').AsInteger;
    RegionCB.Items.Objects[RegionCB.Items.Count-1]:=TObject(pid);
    Table.Next;
  end;
  RegionCB.Sorted:=true;
  if RegionCB.Items.Count>0 then RegionCB.ItemIndex:=0;
  //заполняем список типов
  Table.Close;
  Table.TableName:='MAINTYPES';
  Table.Open;
  ModelCB.Clear;
  while not Table.Eof do begin
    MainTypeCB.Items.Add(Table.FieldByName('DESCR').AsString);
    new(pid);
    pid^:=Table.FieldByName('ID').AsInteger;
    MainTypeCB.Items.Objects[MainTypeCB.Items.Count-1]:=TObject(pid);
    if (MainTypeCB.Items.Count<3)then MainTypeCB.Checked[MainTypeCB.Items.Count-1]:=true;
    Table.Next;
  end;
  Table.Close;
  SetBtn.Down:=true;
  LastDate:=EndOfTheMonth(now);
  LastDateEd.Text:=FormatDateTime('mmmm yyyy',LastDate);
  FirstDate:=StartOfTheMonth(IncYear(now,-1));
  FirstDateEd.Text:=FormatDateTime('mmmm yyyy',FirstDate);
  self.UpdateFilter(self);
  self.UpdateSenterCB;
  FilterPn.Top:=0;
  FirstStart:=true;
  SB.VertScrollBar.Position:=0;
end;

end.

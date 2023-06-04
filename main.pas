unit MAIN;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, Menus,
  StdCtrls, Dialogs, Buttons, Messages, ExtCtrls, ComCtrls, StdActns,
  ActnList, ToolWin, ImgList, DataUnit, DB, ADODB, ActnMan, ActnCtrls,
  XPStyleActnCtrls;

type
  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Window1: TMenuItem;
    Help1: TMenuItem;
    FileExitItem: TMenuItem;
    WindowCascadeItem: TMenuItem;
    WindowTileItem: TMenuItem;
    WindowArrangeItem: TMenuItem;
    HelpAboutItem: TMenuItem;
    WindowMinimizeItem: TMenuItem;
    WindowTileItem2: TMenuItem;
    BottomPanel: TPanel;
    Pages: TPageControl;
    StatusBar: TStatusBar;
    N1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    ActionManager1: TActionManager;
    ActionToolBar1: TActionToolBar;
    MainReport: TAction;
    MonthFoProduction: TAction;
    MonthForCode: TAction;
    WindowClose1: TWindowClose;
    WindowCascade2: TWindowCascade;
    WindowTileHorizontal2: TWindowTileHorizontal;
    WindowTileVertical2: TWindowTileVertical;
    WindowMinimizeAll2: TWindowMinimizeAll;
    WindowArrange1: TWindowArrange;
    Close1: TMenuItem;
    N6: TMenuItem;
    FileExit1: TFileExit;
    N2: TMenuItem;
    procedure AddPage(capt:string;tg,imind:integer);
    procedure PagesChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure OnChangeForm(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FileExit1Hint(var HintStr: string; var CanShow: Boolean);
    procedure HelpContents1Execute(Sender: TObject);
    procedure MainReportExecute(Sender: TObject);
    procedure MonthFoProductionExecute(Sender: TObject);
    procedure MonthForCodeExecute(Sender: TObject);
    procedure N2Click(Sender: TObject);
  private
    { Private declarations }
    function  GetNewChildTAG(clname : shortstring):integer;
    procedure OnCloseChild(var Msg: TMessage); message WM_CLOSECHILD;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses RepStatForRecWin, RepCodeForMonths, RepStatForMonths, About;


procedure TMainForm.FileExit1Hint(var HintStr: string; var CanShow: Boolean);
begin
  Close;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
  i : integer;
begin
  for I := 0 to self.MDIChildCount - 1 do self.MDIChildren[i].Close;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  screen.OnActiveFormChange :=self.OnChangeForm;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  screen.OnActiveFormChange:=nil;
end;

procedure TMainForm.OnChangeForm(Sender: TObject);
var
  i : integer;
begin
  if self.MDIChildCount>0 then begin
    i:=0;
    while (i<Pages.PageCount)and(self.ActiveMDIChild.Tag<>Pages.Pages[i].Tag) do inc(i);
    if (i<Pages.PageCount)and(self.ActiveMDIChild.Tag=Pages.Pages[i].Tag) then Pages.ActivePage:=Pages.Pages[i];
  end;
end;

procedure TMainForm.HelpContents1Execute(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

//------------------------- ÑÎÇÄÀÍÈÅ ÄÎ×ÅÐÍÈÕ ÎÊÎÍ -----------------------------

function  TMainForm.GetNewChildTAG(clname : shortstring):integer;
var
  i,id : integer;
  find : boolean;
begin
  i:=0;
  while(i<MDIChildCount)and(MDIChildren[i].ClassName<>clname) do inc(i);
  if (i<MDIChildCount)and(MDIChildren[i].ClassName=clname)then begin
    MDIChildren[i].BringToFront;
    result:=0;
  end else begin
    id:=1;
    if application.MainForm.MDIChildCount>0 then begin
      find:=true;
      while find do begin
        find:=false;
        i:=0;
        repeat
          if(application.MainForm.MDIChildren[i].Tag=id)then find:=true else inc(i);
        until (i=application.MainForm.MDIChildCount)or(find);
        if find then inc(id);
      end;
    end;
    result:=id;
  end;
end;

procedure TMainForm.MonthFoProductionExecute(Sender: TObject);
var
  form : TRepStatForMonthsForm;
  tg   : integer;
begin
  tg:=self.GetNewChildTAG('');
  if tg>0 then begin
    form := TRepStatForMonthsForm.Create(application);
    form.Tag:=tg;
    Form.WindowState:=wsMaximized;
    AddPage(form.Caption,form.Tag,25);
  end;
end;

procedure TMainForm.MainReportExecute(Sender: TObject);
var
  form : TRepStatForRecForm;
  tg   : integer;
begin
  tg:=self.GetNewChildTAG('');
  if tg>0 then begin
    form := TRepStatForRecForm.Create(application);
    form.Tag:=tg;
    AddPage(form.Caption,form.Tag,25);
  end;
end;

procedure TMainForm.MonthForCodeExecute(Sender: TObject);
var
  form : TRepCodetForMonthsForm;
  tg   : integer;
begin
  tg:=self.GetNewChildTAG('');
  if tg>0 then begin
    form := TRepCodetForMonthsForm.Create(application);
    form.Tag:=tg;
    AddPage(form.Caption,form.Tag,25);
  end;
end;

procedure TMainForm.N2Click(Sender: TObject);
var
  table : TADOQuery;
begin
  table := TADOQuery.Create(self);
  table.Connection:=DataMod.Connection;
  table.SQL.Text := 'SELECT MODELNOTE FROM SERVRECORDS';
  table.Open;
  while not table.Eof do begin
    table.Edit;
    table.FieldByName('MODELNOTE').AsString := Trim(table.FieldByName('MODELNOTE').AsString);
    table.Post;
    table.Next;
  end;
  table.Free  ;
  ShowMessage('Îáðàáîòêà çàâåðøåíà !');
end;

//-------------------------- ÎÊÎÍÍÛÅ ÊÍÎÏÊÈ ------------------------------------

procedure TMainForm.AddPage(capt:string;tg,imind:integer);
var
  tab : TTabSheet;
begin
  tab := TTabSheet.Create(Pages);
  tab.PageControl:=pages;
  tab.Caption:=capt;
  tab.Tag:=tg;
  tab.ImageIndex:=imind;
  if Pages.TabWidth*Pages.PageCount>Pages.ClientWidth then
    Pages.TabWidth:=round((Pages.ClientWidth-10)/pages.PageCount);
  Pages.ActivePage:=tab;
end;

procedure TMainForm.PagesChange(Sender: TObject);
var
  i : integer;
begin
  i:=0;
  while(i<MDIChildCount)and(MDIChildren[i].Tag<>Pages.ActivePage.Tag) do inc(i);
  if (i<MDIChildCount)and(MDIChildren[i].Tag=Pages.ActivePage.Tag) then MDIChildren[i].BringToFront;
end;

procedure TMainForm.OnCloseChild(var Msg: TMessage);
var
  tg,i:integer;
begin
  tg:=msg.wParam;
  i:=0;
  while (i<Pages.PageCount)and(Pages.Pages[i].Tag<>tg) do inc(i);
  if (i<Pages.PageCount)and(Pages.Pages[i].Tag=tg) then begin
    Pages.Pages[i].Destroy;
    if Pages.TabWidth*Pages.PageCount>Pages.ClientWidth then
      Pages.TabWidth:=round((Pages.ClientWidth-10)/pages.PageCount);
  end;
end;


end.

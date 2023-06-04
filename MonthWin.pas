unit MonthWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons;

type
  TMonthDLG = class(TForm)
    Button13: TButton;
    Button14: TButton;
    YearPN: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    BitBtn12: TBitBtn;
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;


function GetMonth(var mn,yr:word):string;

implementation

{$R *.dfm}
uses DateUtils;

var
  newmn,selectyear,newyear : integer;
  monthname : string;

function GetMonth(var mn,yr:word):string;
var
  form: TMonthDLG;
  btn : TButton;
begin
  form:=TMonthDLG.Create(application);
  if (mn>0 )and(mn<13)then begin
    btn:=(form.FindChildControl('BitBtn'+inttostr(mn))as TButton);
    btn.Font.Style:=[fsBold];
    btn.Font.Color:=clRED;
    form.Repaint;
    btn.Focused;
  end;
  selectyear:=yr;
  if selectyear=0 then newyear:=YearOf(now) else begin
    form.YearPN.Font.Color:=clRed;
    newyear:=selectyear;
  end;
  form.YearPN.Caption:=IntToStr(newyear);
  if form.ShowModal=mrOk then begin
    mn:=newmn;
    yr:=newyear;
    result:=monthname+' '+IntTostr(newyear);
  end else result:='';
  form.Free;
end;

procedure TMonthDLG.Button13Click(Sender: TObject);
begin
  dec(newyear);
  if newyear=selectyear then YearPN.Font.Color:=clRed else YearPN.Font.Color:=clBlack;
  YearPN.Caption:=IntToStr(newyear);
end;

procedure TMonthDLG.Button14Click(Sender: TObject);
begin
  inc(newyear);
  if newyear=selectyear then YearPN.Font.Color:=clRed else YearPN.Font.Color:=clBlack;
  YearPN.Caption:=IntToStr(newyear);
end;

procedure TMonthDLG.Button1Click(Sender: TObject);
var
  str:string;
begin
  str:=(sender as TControl).Name;
  monthname:=(sender as TBitBtn).Caption;
  delete(str,1,length('BitBtn'));
  newmn:=strtointdef(str,0);
  if newmn=0 then Abort;
  self.ModalResult:=mrOk;
end;

end.

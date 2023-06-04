unit MonthGridType;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
  Dialogs, ExtCtrls, Grids, StdCtrls, DataUnit;

type
  TGridRange = record
    left,Top,Right,Bottom : integer;
    Text                  : string[255];
    Align                 : cardinal;
    BrushColor, PenColor  : TColor;
    Font                  : TFont;
  end;
  TRangeList = array of TGridRange;

  TMyGrid = class(TStringGrid)
  protected
    procedure Paint; override;
  private
    { Private declarations }
    FRanges        : TRangeList;
    FFixedPenColor : TColor;
    FFixedFont     : TFont;
    FTimePeriods   : TTimePeriod;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent);override;
    destructor  Destroy;override;
    function    AddRange(X1,Y1,X2,Y2:integer;txt:string):integer;
    procedure   ClearRanges;
    procedure   ClearRows;
    procedure   RepaintHeader(firstdate,lastdate : TDateTime; mode : integer);
    procedure   ChangeText(X1,Y1:integer; txt : string);
    property    TimePeriod : TTimePeriod read FTimePeriods;
    procedure   DrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure   MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  end;

implementation

//--------- базовый класс таблицы с многостраничным заголовком -----------------

constructor TMyGrid.Create(AOwner: TComponent);
begin
  inherited;
  SetLength(self.FRanges,0);
  self.RowCount:=3;
  self.FixedRows:=2;
  self.ColCount:=3;
  self.FixedCols:=2;
  self.Ctl3D:=false;
  self.DefaultColWidth:=20;
  self.Options:= self.Options+[goColSizing];
  self.FixedColor:=clWhite;
  self.FFixedFont:=self.Font;
  self.FFixedPenColor:=clBlack;
  self.FTimePeriods:=TTimePeriod.Create;
  self.OnDrawCell:=self.DrawCell;
  self.OnMouseDown:=self.MouseDown;
end;

destructor TMyGrid.Destroy;
begin
  SetLength(self.FRanges,0);
  self.FRanges:=nil;
  self.FTimePeriods.Free;
  inherited;
end;

procedure TMyGrid.ClearRows;
var
  i : integer;
begin
  for I := self.FixedRows to self.RowCount - 1 do self.Rows[i].Clear;
  self.RowCount:=self.FixedRows+1;
end;

function TMyGrid.AddRange(X1,Y1,X2,Y2:integer;txt:string):integer;
begin
  SetLength(self.FRanges,high(self.FRanges)+2);
  self.FRanges[high(self.FRanges)].left:=X1;
  self.FRanges[high(self.FRanges)].Right:=X2;
  self.FRanges[high(self.FRanges)].Top:=Y1;
  self.FRanges[high(self.FRanges)].Bottom:=Y2;
  self.FRanges[high(self.FRanges)].Text:=txt;
  self.Cells[X1,Y1]:=txt;
  result:=high(self.FRanges)
end;

procedure TMyGrid.ClearRanges;
begin
  setlength(self.FRanges,0);
end;

procedure TMyGrid.ChangeText(X1,Y1:integer; txt : string);
var
  i : integer;
begin
  i:=0;
  while(i<=high(self.FRanges))and(not((self.FRanges[i].left=X1)and(self.FRanges[i].Top=Y1)))do inc(i);
  if(i<=high(self.FRanges))and(self.FRanges[i].left=X1)and(self.FRanges[i].Top=Y1)then self.FRanges[i].Text:=txt;
end;

procedure TMyGrid.RepaintHeader(firstdate,lastdate : TDateTime; mode : integer);
var
  col,m,w : integer;
  mn      : TDateTime;
  txt     : string;
begin
  self.FTimePeriods.SetDates(firstdate,lastdate,mode);
  self.ClearRanges;
  self.ColCount:=self.FixedCols+self.FTimePeriods.PeriodsCount+1;
  if self.FixedCols>0 then self.AddRange(0,0,self.FixedCols-1,1,'Модель');
  self.AddRange(self.FixedCols,0,self.FixedCols,1,'Всего');
  col:=self.FixedCols+1;
  for m := 0 to self.FTimePeriods.Count - 1 do begin
    mn:=EncodeDate(self.FTimePeriods.Month[m].year,self.FTimePeriods.Month[m].month,1);
    self.AddRange(col,0,col+(high(self.FTimePeriods.Month[m].dates)+1)-1,0,
      FormatDateTime('mmmm yyyy',mn));
    col:=col+(high(self.FTimePeriods.Month[m].dates)+1);
  end;
  w:=0;
  for m := 0 to self.FTimePeriods.PeriodsCount - 1 do begin
    col:=self.FixedCols+1+m;
    txt:=FormatDateTime('dd.mm',self.FTimePeriods.Period[m].first)
      +' - '+FormatDateTime('dd.mm',self.FTimePeriods.Period[m].last);
    self.AddRange(col,1,col,1,txt);
    if round(self.Canvas.TextWidth(txt)*1.3)>w then w:=round(self.Canvas.TextWidth(txt)*1.3);
  end;
  for m := self.FixedCols to self.ColCount-1 do if self.ColWidths[m]<w then self.ColWidths[m]:=w;
 self.Paint;
end;

procedure TMyGrid.Paint;
var
  rct   : TRect;
  i     : integer;
  Rng   : TGridRange;
  txt   : string;
  candraw : boolean;
  fnt   : TFont;
  bclr  : TColor;
  pn    : TPen;

procedure VerticalForFixed;
var
  j : integer;
begin
  //расчет вертикальных координат для фиксированных строк
  rct.Top:=0;
  for j:=0 to rng.Top-1 do rct.Top:=rct.Top+self.RowHeights[j]+1;
  rct.Bottom:=rct.Top;
  for j:=rng.Top to rng.Bottom do rct.Bottom:=rct.Bottom+self.RowHeights[j]+1;
end;

procedure HorisontalForFixed;
var
  j : integer;
begin
  //расчет горизонтальных координат для фиксированных строк
  rct.Left:=0;
  for j:=0 to rng.left-1 do rct.Left:=rct.Left+self.ColWidths[j]+1;
  rct.Right:=rct.Left;
  for j:=rng.left to rng.Right do rct.Right:=rct.Right+self.ColWidths[j]+1;
end;

procedure HorisontalForMoved;
var
  j : integer;
begin
  //вычисление левой позиции зависит от того, где находится начало прямоугольника
  //слева от ColLeft справа
  rct.left:=0;
  if rng.left>=leftcol then begin
    for j:=0 to self.FixedCols-1 do rct.left:=rct.left+self.ColWidths[j]+1;
    for j:=leftcol to rng.left-1 do rct.left:=rct.left+self.ColWidths[j]+1;
  end;
  if rng.left<leftcol then begin
    for j:=0 to self.FixedCols-1 do rct.left:=rct.left+self.ColWidths[j]+1;
    for j:=rng.left to leftcol-1 do rct.left:=rct.left-self.ColWidths[j]-1;
  end;
  //вычислем кординаты правого нижнего угла относительно левого
  rct.Right:=rct.Left;
  for j := rng.left to rng.Right do rct.Right:=rct.Right+self.ColWidths[j]+1;
end;

procedure VerticalForMoved;
var
  j : integer;
begin
  //вычисление верхней позиции зависит от того, где находится начало прямоугольника
  //сверху от TopRow или снизу
  rct.Top:=0;
  if rng.Top>=toprow then begin
    for j:=0 to self.FixedRows-1 do rct.top:=rct.top+self.RowHeights[j]+1;
    for j:=toprow to rng.top-1 do rct.top:=rct.top+self.RowHeights[j]+1;
  end;
  if rng.top<toprow then begin
    for j:=0 to self.FixedRows-1 do rct.top:=rct.top+self.RowHeights[j]+1;
    for j:=rng.top to toprow-1 do rct.top:=rct.top-self.RowHeights[j]-1;
  end;
  //вычислем кординаты правого нижнего угла относительно левого
  rct.Bottom:=rct.Top;
  for j := rng.Top to rng.Bottom do rct.Bottom:=rct.Bottom+self.RowHeights[j]+1;
end;

begin
  inherited;
  fnt:=self.Canvas.Font;
  bclr:=self.Canvas.Brush.Color;
  pn:=self.Canvas.Pen;
  //рисуем диапазоны в подвшижных ячейках диапазоны
  for i := low(self.FRanges) to high(self.FRanges) do begin
    Rng:=self.FRanges[i];
    if (rng.left>=self.FixedCols)and(rng.Top>=self.FixedRows)and
    (rng.Bottom>=TopRow)and(self.CellRect(rng.left,rng.Top).Top<self.ClientRect.Bottom)and
    (rng.Right>=LeftCol)and(self.CellRect(rng.left,rng.Top).Left<self.ClientRect.Right) then begin
      VerticalForMoved;
      HorisontalForMoved;
      if rng.Font<>nil then self.Canvas.Font:=rng.Font;
      if rng.BrushColor<>0 then self.Canvas.Brush.Color:=rng.BrushColor;
      if rng.pencolor<>0 then self.Canvas.Pen.Color:=rng.pencolor;
      self.Canvas.Rectangle(rct);
      inc(rct.Top,2);inc(rct.Left,2);
      dec(rct.Bottom,2); dec(rct.Right,2);
      txt:=rng.Text;
      DrawText(canvas.Handle,pchar(txt),Length(txt),rct,DT_CENTER);
    end;
  end;
  //рисуем фиксированные строки и столбцы
  for i := low(self.FRanges) to high(self.FRanges) do begin
    Rng:=self.FRanges[i];
    candraw:=false;
    if (rng.left>=0)and(rng.Right<self.FixedCols)and
    (rng.Bottom>=TopRow)and(self.CellRect(rng.left,rng.Top).Top<self.ClientRect.Bottom)then begin
      VerticalForMoved;
      HorisontalForFixed;
      dec(rct.Left);
      candraw:=true;
    end;
    if (rng.Top>=0)and(rng.Bottom<self.FixedRows)and
    (rng.Right>=LeftCol)and(self.CellRect(rng.left,rng.Top).Left<self.ClientRect.Right) then begin
      VerticalForFixed;
      HorisontalForMoved;
      dec(rct.top);
      candraw:=true;
    end;
    if candraw then begin
      self.Canvas.Pen.Color:=self.FFixedPenColor;
      self.Canvas.Font:=self.FFixedFont;
      self.Canvas.Rectangle(rct);
      //вывод текста
      inc(rct.Top,2);inc(rct.Left,2);
      dec(rct.Bottom,2); dec(rct.Right,2);
      txt:=rng.Text;
      DrawText(canvas.Handle,pchar(txt),Length(txt),rct,DT_CENTER);
    end;
  end;
  //рисуем неподвижные ячейки правого верхнего угла
  for i := low(self.FRanges) to high(self.FRanges) do begin
    Rng:=self.FRanges[i];
    if (rng.left>=0)and(rng.Right<self.FixedCols)and(rng.Top>=0)and(rng.Bottom<self.FixedRows)then begin
      HorisontalForFixed;
      VerticalForFixed;
      dec(rct.top);
      dec(rct.Left);
      self.Canvas.Pen.Color:=self.FFixedPenColor;
      self.Canvas.Font:=self.FFixedFont;
      self.Canvas.Rectangle(rct);
      //вывод текста
      inc(rct.Top,2);inc(rct.Left,2);
      dec(rct.Bottom,2); dec(rct.Right,2);
      txt:=rng.Text;
      DrawText(canvas.Handle,pchar(txt),Length(txt),rct,DT_CENTER);
    end;
  end;
  self.Canvas.Font:=fnt;
  self.Canvas.Brush.Color:=bclr;
  self.Canvas.Pen:=pn;
end;

procedure TMyGrid.DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  str : string;
  rct : TRect;
  fl  : cardinal;
  h   : integer;
begin
  if ARow>=self.FixedRows then begin
    rct:=rect;
    str:=self.Cells[ACol,ARow];
    self.Canvas.FillRect(rct);
    inc(rct.Top,2);inc(rct.Left,2);
    dec(rct.Bottom,2); dec(rct.Right,2);
    if ACol=1 then fl:=DT_WORDBREAK else fl:=DT_CENTER or DT_VCENTER;
    h:=DrawText(canvas.Handle,pchar(str),Length(str),rct,fl);
    if (ACol=1)and((h+8)>self.RowHeights[ARow]) then self.RowHeights[ARow]:=h+8;
  end;
end;

procedure TMyGrid.MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  crd : TGridCoord;
begin
  //левая кнопка мыши - выделение
  if Button=mbLeft then begin
    crd:=MouseCoord(X,Y);
    if (crd.Y>=0)and(crd.Y<self.FixedCols)and(crd.X>=0)and(crd.X<self.FixedRows) then
      Selection:=TGridRect(rect(0,0,ColCount-1,RowCount-1));
    //шелчок по заголовку столбца - левая кнопка выдлеение столбца
    if (crd.Y>=0)and(crd.Y<self.FixedCols)and(crd.X>=FixedCols) then
      Selection:=TGridRect(rect(crd.X,0,crd.X,RowCount-1));
    //шелчок по заголовку строки левая кнопка мыши - выделение строки
    if (crd.X=0)and(crd.Y>=FixedRows) then
      Selection:=TGridRect(rect(0,crd.Y,ColCount-1,crd.Y));
  end;
end;


end.

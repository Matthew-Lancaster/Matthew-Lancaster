program labell;

uses    GRAPH,  KEY,  TEXTS,  PLOT;








var

size   : byte;

name   : string;







procedure plottitle;
var
xsize, ysize, xoff, yoff    : word;
x1, x2, y1, y2, L, L1, L2   : word;
midx, midy                  : word;
Lx , Ly                     : real;
 begin
 loadpaper(A4);
 if CR then
  begin
  openplot(vertical);
  selectpen;
  Ly:= 0.15*size;     { cm }
  Lx:= Ly/1.25;
  write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2);
  write(Lst, 'LO 2');
  xsize:= plotunits(Lx*10*1.5);
  ysize:= plotunits(Ly*10);
  papermidpoints(vertical, A4, midx, midy);
  L1:= round(9.25*xsize) - plotunits(Lx*0.5);      { ROY JENNER }
  xoff:= midx - (L1 div 2);
  yoff:= midy + round(ysize*2);
  write(Lst, 'PA', xoff, ' ', yoff);
  write(Lst, 'LB', 'ROY', #3);
  xoff:= xoff + round(3.25*xsize);
  write(Lst, 'PA', xoff, ' ', yoff);
  write(Lst, 'LB', 'JENNER', #3);
  L2:= (length(name)*xsize) - plotunits(Lx*0.5);   { TITLE }
  xoff:= midx- (L2 div 2);
  yoff:= midy - round(ysize*2);
  write(Lst, 'PA', xoff, ' ', yoff);
  write(Lst, 'LB', name, #3);
  if (L2>= L1) then
   L:= L2
   else
   L:= L1;
  x1:= midx - (L div 2) - round(ysize*2.5);
  x2:= midx + (L div 2) + round(ysize*2);
  y1:= midy + (ysize*5);
  y2:= midy - (ysize*5);
  plotbox(x1, y1, x2, y2, 1, 1, 0);
  closeplot;
  end
  else resetkeys;
 end;   { of procedure plottitle }









procedure drawsize;
var
y1, y2  : word;
 begin
 setfillstyle(1,1);
 bar(310, 300, 330, 360);
 rectangle(310, 300, 330, 360);
 line(310, 320, 330, 320);
 line(310, 340, 330, 340);
 setfillstyle(1,0);
 y1:= 360 - ((size-1)*20);
 y2:= y1-20;
 bar(310, y1, 330, y2);
 outtextxy(306, 370, 'SIZE');
 end;   { of drawsize }









procedure readtitle;
 begin
 setcolor(black);
 drawsize;
 outtextxy(300, 250, 'TITLE : ');
 ord(name[0]):= 0;
 repeat
  readstring(300 + (9*8), 250, black, 50, true, name);
  presskey;
  if CurUp and (size< 3) then inc(size);
  if CurDn and (size> 1) then dec(size);
  if CurUp or CurDn then drawsize;
 until CR or esc;
 erasetextcursor(white);
 if CR then plottitle;
 end;   { of procedure readtitle }







begin


size:= 1;

repeat
 readtitle;
until esc;


closegraph;


end.   { of program labell }
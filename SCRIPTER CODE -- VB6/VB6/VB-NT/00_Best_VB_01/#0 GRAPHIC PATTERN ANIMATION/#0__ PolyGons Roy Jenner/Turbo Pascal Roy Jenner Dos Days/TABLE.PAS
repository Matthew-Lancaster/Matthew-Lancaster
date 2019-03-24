program TABLE;  { plots table }

uses   GRAPH,  KEY,  TEXTS,  DISPLAY,  PLOT;









var

icon              : string[8];

ratio             : real;








procedure tablescreen;
var
x1, x2, y1, y2, L, midx : word;
 begin
 disptext('table', screenx div 2, screeny*15 div 100, black, 6);
 L:= screenx*50 div 100;
 midx:= screenx div 2;
 x1:= midx - L div 2;
 x2:= midx + L div 2;
 y1:=  screeny*35 div 100;
 y2:= y1+24;
 displaybox(black, x1, y1, x2, y2, 20, 2);
 settextstyle(0,0,1);
 outtextxy(x1+3, y1+2,  'x');
 outtextxy(x1+3, y1+14, 'y');
 end;   { of procedure tablescreen }








procedure labelbox(x, y : word);
 begin
 writeln( Lst, 'PA', x+160, ' ', y+440);
 writeln( Lst, 'LO 1', 'LB', 'x', #3);
 writeln( Lst, 'PA', x+160, ' ', y+120);
 writeln( Lst, 'LO 1', 'LB', 'y', #3);
 end;   { of procedure labelbox }








procedure plottable;
var
W, H, L        : word;
x1, y1         : word;
midx, midy     : word;
 begin
 loadpaper(A4);
 if CR then
  begin
  openplot(horizontal);
  W:= 200;   { mm }
  H:= 16;
  W:= plotunits(W);
  H:= plotunits(H);
  papermidpoints(horizontal, A4, midx, midy);
  x1:= midx - (W div 2);
  y1:= 160*40;
  repeat
   plotbox(x1, y1, x1+W, y1+H, 20, 2, 0);
   labelbox(x1, y1);
   dec(y1,round(H*1.5));
  until y1 < 800;
  L:= round(length(icon)*0.187*400*1.5);
  x1:= midx - (W div 4) - (L div 2);
  y1:= 800;
  writeln( Lst, 'PA', x1, ' ', y1);
  writeln( Lst, 'LO 1', 'LB', icon, #3);
  str(ratio:2:2, icon);
  x1:= midx + (W div 4) - (L div 2);
  writeln( Lst, 'PA', x1, ' ', y1);
  writeln( Lst, 'LO 1', 'LB', icon, #3);
  closeplot;
  end
  else resetkeys;
 end;   { of procedure plottable }








procedure readtabledata;
var
xoff, yoff, xmid, line   : word;
X, Y, L                  : word;
code, code1              : integer;
data                     : array[0..2] of string[8];
gridsize                 : real;
 begin
 line:= 0;
 xmid:= screenx div 2;
 disptext('   ICON  :         ',  xmid, screeny*65 div 100, black, 1);
 disptext('X size :     mm',  xmid, screeny*80 div 100, black, 1);
 disptext('Y size :     mm',  xmid, screeny*85 div 100, black, 1);
 xoff:= xmid - (15*8 div 2) +68;
 ord(data[0][0]):= 0;
 ord(data[1][0]):= 0;
 ord(data[2][0]):= 0;
 repeat
  if line = 0 then
   yoff:= (screeny*65 div 100) - 4;
  if line = 1 then
   yoff:= (screeny*80 div 100) - 4;
  if line = 2 then
   yoff:= (screeny*85 div 100) - 4;
  if line= 0 then L:= 8 else L:= 3;
  readstring(xoff, yoff, black, L, true, data[line]);
  presskey;
  val(data[1],  gridsize, code);
  X:= round(gridsize);
  val(data[2],  gridsize, code1);
  Y:= round(gridsize);
  if CR and (line= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 1;
   resetkeys;
   end;
  if CR and (line= 1) and (code= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 2;
   resetkeys;
   end;
  if CurDn and (line= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 1;
   resetkeys;
   end;
  if CurDn and (line= 1) and (code= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 2;
   end;
  if CurUp and (line= 1) and (code= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 0;
   end;
  if CurUp and (line= 2) and (code1= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 1;
   end;
 until (CR and (line= 2) and (code= 0) and (code1=0) ) or esc;
 erasetextcursor(bkgcol);
 icon:= data[0];
 if (x<> 0 ) then ratio:= y/x;
 if CR then plottable;
 end;   { of procedure readtabledata }







begin


tablescreen;
plotquality:= highquality;
repeat
 readtabledata;
until esc;
closegraph;



end.   { of program TABLE }
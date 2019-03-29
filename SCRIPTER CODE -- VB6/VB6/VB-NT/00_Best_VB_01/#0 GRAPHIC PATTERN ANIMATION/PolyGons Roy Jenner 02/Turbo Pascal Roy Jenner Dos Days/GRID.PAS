program grid;

uses   GRAPH,  KEY,  TEXTS,   DISPLAY,  PLOT;








var




direction        : byte;

papersize        : byte;

X, Y             : word;

W                : word;

xparts, yparts   : word;

gridx, gridy     : longint;

gx1, gy1         : longint;








procedure gridscreen;
var
x1, x2, y1, y2   : word;
 begin
 disptext('grid', screenx div 2, screeny*20 div 100, black, 6);
 x1:= screenx*40 div 100;
 x2:= screenx*60 div 100;
 y1:= screeny*40 div 100;
 y2:= screeny*60 div 100;
 displaybox(black, x1, y1, x2, y2, 15, 15);
 end;   { of procedure gridscreen }







procedure labelgrid;
var
i, xplot, yplot : word;
 begin
 writeln(Lst, 'SI', '0.14', ',', '0.2', ';');
 if XPARTS >= 10 then
  begin
  yplot:= gy1;
  writeln(Lst, 'LO 16');
  for i:= 0 to (XPARTS div 10) do   { bottem axis }
   begin
   xplot:= gx1 + ( (i*gridx*10) div XPARTS );
   writeln( Lst, 'PA', xplot:6, yplot:6);
   writeln( Lst, 'PD', xplot:6, yplot-100:6, 'PU');
   xplot:= xplot + W;
   writeln( Lst, 'PA', xplot:6, yplot:6);
   writeln( Lst, 'PD', xplot:6, yplot-100:6, 'PU');
   writeln(Lst, 'LB', (i*10), chr(3) );
   end;
  writeln(Lst, 'LO 14');
  for i:= 0 to (XPARTS div 10) do   { top axis }
   begin
   xplot:= gx1 + ( (i*gridx*10) div XPARTS );
   writeln( Lst, 'PA', xplot:6, yplot+gridy:6);
   writeln( Lst, 'PD', xplot:6, yplot+gridy+100:6, 'PU');
   xplot:= xplot + W;
   writeln( Lst, 'PA', xplot:6, yplot+gridy:6);
   writeln( Lst, 'PD', xplot:6, yplot+gridy+100:6, 'PU');
   writeln(Lst, 'LB', (i*10), chr(3) );
   end;
  end;
if YPARTS >= 10 then
  begin
  xplot:= gx1;
  writeln(Lst, 'LO 18');
  for i:= 0 to (YPARTS div 10) do    { left axis }
   begin
   yplot:= gy1 + ( (i*gridy*10) div YPARTS );
   writeln( Lst, 'PA', xplot:6, yplot:6);
   writeln( Lst, 'PD', xplot-100:6, yplot:6, 'PU');
   yplot:= yplot + W;
   writeln( Lst, 'PA', xplot:6, yplot:6);
   writeln( Lst, 'PD', xplot-100:6, yplot:6, 'PU');
   writeln(Lst, 'LB', (i*10), chr(3) );
   end;
  writeln(Lst, 'LO 12');
  for i:= 0 to (YPARTS div 10) do    { right axis }
   begin
   yplot:= gy1 + ( (i*gridy*10) div YPARTS );
   writeln( Lst, 'PA', xplot+gridx:6, yplot:6);
   writeln( Lst, 'PD', xplot+gridx+100:6, yplot:6, 'PU');
   yplot:= yplot + W;
   writeln( Lst, 'PA', xplot+gridx:6, yplot:6);
   writeln( Lst, 'PD', xplot+gridx+100:6, yplot:6, 'PU');
   writeln(Lst, 'LB', (i*10), chr(3) );
   end;
  end;
 end;   { of procedure labelgrid }









procedure plotgrid;
var
penwidth        : real;
i, xplot, yplot : word;
midx, midy      : word;
 begin
 plotquality:= highquality;
 loadpaper(papersize);
 if CR then
  begin
  papermidpoints(direction, papersize, midx, midy);
  penwidth:= 0.35;     { mm }
  W:= plotunits(penwidth*0.7);
  gridx:= plotunits(X);
  gridy:= plotunits(Y);
  gx1:= midx - (gridx div 2);
  gy1:= midy - (gridy div 2);
  openplot(direction);
  plotbox(gx1, gy1, gx1+gridx, gy1+gridy, xparts, yparts, 0);
  if XPARTS >= 10 then
   begin
   yplot:= gy1;
   for i:= 0 to (XPARTS div 10) do
    begin
    xplot:= gx1 + ( (i*gridx*10) div XPARTS ) + W;
    writeln( Lst, 'PA', xplot:6, yplot:6);
    writeln( Lst, 'PD', xplot:6, yplot+gridy:6, 'PU');
    if i< (XPARTS div 10) then
     begin
     xplot:= xplot + ( ((gridx*10) div XPARTS) div 2);
     writeln( Lst, 'PA', xplot:6, yplot:6);
     writeln( Lst, 'PD', xplot:6, yplot+gridy:6, 'PU');
     end;
    end;
   end;
  if YPARTS>= 10 then
   begin
   xplot:= gx1;
   for i:= 0 to (YPARTS div 10) do
    begin
    yplot:= gy1 + ( (i*gridy*10) div YPARTS ) + W;
    writeln( Lst, 'PA', xplot:6, yplot:6);
    writeln( Lst, 'PD', xplot+gridx:6, yplot:6, 'PU');
    if i< (YPARTS div 10) then
     begin
     yplot:= yplot + ( ((gridy*10) div YPARTS) div 2);
     writeln( Lst, 'PA', xplot:6, yplot:6);
     writeln( Lst, 'PD', xplot+gridx:6, yplot:6, 'PU');
     end;
    end;
   end;
  labelgrid;
  closeplot;
  end
  else resetkeys;
 end;   { of procedure plotgrid }








procedure sizepaper(x, y : word; var direction, papersize : byte);
var
xrange, yrange  : word;
 begin
 if (x> y) then
  direction:= horizontal
  else
  direction:= vertical;
 paperranges(direction, A4, xrange, yrange);
 if (x <= xrange) and ( y <= yrange) then
  papersize:= A4
  else
   begin
   paperranges(direction, A3, xrange, yrange);
   if (x <= xrange) and ( y <= yrange) then
   papersize:= A3
   else
    begin
    papersize:= 2;
    opentextwindow(white, 12, 2);
    outtextxy(8, 10, 'plot too big');
    repeat
     presskey;
    until esc;
    closetextwindow;
    resetkeys;
    end;
   end;
 end;   { of procedure sizepaper }








procedure readgriddata;
var
xoff, yoff, xmid, line   : word;
code, code1              : integer;
size                     : array[0..1] of string[3];
gridsize                 : real;
 begin
 line:= 0;
 xmid:= screenx div 2;
 disptext('X size :      mm ', xmid, screeny*75 div 100, black, 1);
 disptext('Y size :      mm ', xmid, screeny*80 div 100, black, 1);
 xoff:= xmid - (15*8 div 2) +68;
 ord(size[0][0]):= 0;
 ord(size[1][0]):= 0;
 repeat
  if line = 0 then
   yoff:= (screeny*75 div 100) - 4;
  if line = 1 then
   yoff:= (screeny*80 div 100) - 4;
  readstring(xoff, yoff, black, 3, true, size[line]);
  presskey;
  val(size[0],  gridsize, code);
  X:= round(gridsize);
  val(size[1],  gridsize, code1);
  Y:= round(gridsize);
  if CR and (line= 0) and (code= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 1;
   resetkeys;
   end;
  if CurDn and (line= 0) and (code= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 1;
   end;
  if CurUp and (line= 1) and (code1= 0) then
   begin
   erasetextcursor(bkgcol);
   line:= 0;
   end;
 until (CR and (line= 1) and (code= 0) and (code1=0)) or esc;
 erasetextcursor(bkgcol);
 sizepaper(X, Y, direction, papersize);
 if (CR and (papersize< 2)) then plotgrid;
end;   { of procedure readgriddata }









begin

xparts:= 91;
yparts:= 91;
gridscreen;
repeat
 readgriddata;
until esc;
closegraph;

end.   { of program grid }
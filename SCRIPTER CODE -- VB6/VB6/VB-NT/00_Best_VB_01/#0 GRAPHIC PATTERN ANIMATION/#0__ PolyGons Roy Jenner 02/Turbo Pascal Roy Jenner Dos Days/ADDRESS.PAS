program address;

uses    GRAPH,  KEY,  TEXTS,  PLOT;








var

xsize, ysize  : word;

Lx , Ly       : real;








procedure plotlabel(x, y : word; s : string);
 begin
 write(Lst, 'PA', x, ' ', y);
 write(Lst, 'LB', s, #3);
 end;   { of procedure plotlabel }









procedure plotaddress;
var
midx, midy  : word;
x, y        : word;
 begin
 loadpaper(A4);
 if CR then
  begin
  openplot(vertical);
  selectpen;
  Ly:= 0.25;     { cm }
  Lx:= Ly/1.5;
  write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2);
  write(Lst, 'LO 2');
  xsize:= plotunits(Lx*10*1.5);
  ysize:= plotunits(Ly*10);
  papermidpoints(vertical, A4, midx, midy);
  x:= midx + plotunits(50);
  y:= midy + plotunits(130);
  plotlabel(x, y, 'Roy Jenner');
  y:= y - round(ysize*2);
  plotlabel(x, y, 'Flat 4');
  y:= y - round(ysize*2);
  plotlabel(x, y, '34 Lansdowne Place');
  y:= y - round(ysize*2);
  plotlabel(x, y, 'Hove BN3 1HH');
  y:= y - round(ysize*3);
  plotlabel(x, y, 'Tel 01273 738307');
  closeplot;
  end
  else resetkeys;
 end;   { of procedure plotaddress }









procedure displayaddress;
var
x, y  : word;
 begin
 x:= 480; y:= 20;
 outtextxy(x, y,  'Roy Jenner');
 outtextxy(x, y+12,  'Flat4');
 outtextxy(x, y+24, '34 Lansdowne Place');
 outtextxy(x, y+36, 'Hove BN3 1HH');
 outtextxy(x, y+54, 'Tel 01273 738307');
 repeat
  presskey;
  if AltP then plotaddress;
 until esc;
 end;   { of procedure displayaddress }







begin



displayaddress;


closegraph;


end.   { of program doorbell }
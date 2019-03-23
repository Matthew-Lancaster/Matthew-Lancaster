program prscale;


uses  GRAPH,  KEY,  TEXTS,  DISPLAY,  PLOT;








var

G                  : array[0..500] of real;

min, mean, max, n  : word;

midx, midy         : word;

xsize, ysize       : word;

x1, y1             : word;









procedure initialise;
 begin
 plotarea:= 100;
 plotquality:= highquality;
 min:= round(0.2*100);
 mean:= round(0.4*100);
 max:= round(2.0*100);   { xysamples for plotter are 225 max }
 end;   { of procedure initialise }









procedure makegraph;
var
xsize, ysize    : word;
xscale, yscale  : real;
xoff, yoff      : word;
x1, y1, x2, y2  : word;
x, y            : real;
s               : string[4];
 begin
 ysize:= 200;
 xsize:= 200*5 div 4;
 xscale:= xsize/(max-min-1);
 xscale:= xscale*(max-min)/max;
 yscale:= ysize/0.8;
 xoff:= (screenx - xsize) div 2;
 yoff:= (screeny - ysize) div 2;
 displaybox(black, xoff, yoff, xoff+xsize, yoff+ysize, 5, 4);
 disptext('PATRANGE SCALE', xoff+round(xsize/2), yoff-40, black, 1);
 x:= 0;
 repeat
  str(x:3:1, s);
  disptext(s, xoff+round(x*xsize/2)+2, yoff+ysize+16, black, 1);
  x:= x + 0.4;
 until x > 2;
 disptext('0.1969/A', xoff+round(xsize/2), yoff+ysize+40, black, 1);
 y:= 0.2;
 repeat
  str(y:4:2, s);
  disptext(s, xoff-24, yoff+ysize-round((y-0.2)*ysize/(1.0-0.2)), black, 1);
  y:= y+0.2;
 until y > 1.0;
 disptext('pr', xoff-75, yoff+round(ysize/2), black, 1);
 G[min]:= 0;
 setlinestyle(0, 0, 3);
 setviewport(xoff, yoff, xoff+xsize, yoff+ysize, clipon);
 for n:= min to max do
  begin
  if (n< mean) then
   G[n-min]:= 0.2 + (n-min)*0.2/(mean-min)
   else
   G[n-min]:= 0.4 + ln(n/mean)*0.6/ln(5);
   end;
 for n:= min to max-1 do
  begin
  x1:= round(n*xscale);
  y1:= round((1.0 - G[n-min])*yscale);
  x2:= round((n+1)*xscale);
  y2:= round((1.0 - G[n-min+1])*yscale);
  line(x1, y1, x2, y2);
 end;
 setlinestyle(0, 0, 1);
 end;   { of procedure makegraph }









procedure labelscale;
var
xoff, yoff      : word;
Lx, Ly, x, y    : real;
s               : string[4 ];
 begin
 resetplotterwindow;
 Ly:= plotarea*0.25/100;
 Lx:= Ly/1.25;
 write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2, 'LO 15');

 xoff:= midx;
 yoff:= y1 + ysize + round(ysize/8);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'PATRANGE SCALE', #3);

 Ly:= plotarea*0.2/100;
 Lx:= Ly/1.25;
 write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2);

 yoff:= y1 - round(ysize/8);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', '0.1969/A', #3);


 xoff:= x1- round(xsize/5);
 yoff:= midy;
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'pr', #3);

 xoff:= x1;
 yoff:= y1 - plotunits(Lx*10);
 x:= 0;
 repeat
  write(Lst, 'PA', xoff, ' ', yoff);
  str(x:3:1, s);
  write(Lst, 'LO 16', 'LB', s, #3);
  x:= x+0.4;
  xoff:= x1 + round(xsize*x/2);
 until x> 2;
 xoff:= x1 - plotunits(Lx*10);
 yoff:= y1;
 y:= 0.2;
 repeat
  write(Lst, 'PA', xoff, ' ', yoff);
  str(y:4:2, s);
  write(Lst, 'LO 18', 'LB', s, #3);
  y:= y+0.2;
  yoff:= y1 + round(ysize*(y-0.2)/0.8);
  until y> 1.0;
 end;   { of procedure labelscale }









procedure plotgraph;
var
xscale, yscale   : real;
L, LW            : word;
 begin
 plotquality:= highquality;
 selectplotsize(1.0, papersize, plotsize);
 if CR then loadpaper(papersize);
 if CR then
  begin
  papermidpoints(horizontal, papersize, midx, midy);
  openplot(direction);
  selectpen;
  ysize:= plotunits(plotsize);
  xsize:= ysize*5 div 4;
  x1:= midx - (xsize div 2);
  y1:= midy - (ysize div 2);
  xscale:= xsize/round(max-min);
  xscale:= xscale*(max-min)/max;
  yscale:= ysize/0.8;
  L:= max - min;
  for n:= 0 to L do
   begin
   xy[n,0]:= x1 + round(min*xscale) + round(n*xscale);
   xy[n,1]:= y1 + round(G[n]*yscale) - round(0.2*yscale);
   end;
  LW:= round((plotarea/100)*35);
  plotbox(x1, y1, x1+xsize, y1+ysize, 5, 4, 0);
  setplotterwindow(x1, y1, x1+xsize, y1+ysize);
  plotline(LW, L);
  labelscale;
  closeplot;
  end;
 end;   { of procedure plotgraph }








procedure editgraph;
 begin
 makegraph;
 repeat
  if AltP then plotgraph;
  presskey;
 until esc;
 end;   { of procedure editgraph }









begin

initialise;

editgraph;

closegraph;


end.   { of program SCALE }

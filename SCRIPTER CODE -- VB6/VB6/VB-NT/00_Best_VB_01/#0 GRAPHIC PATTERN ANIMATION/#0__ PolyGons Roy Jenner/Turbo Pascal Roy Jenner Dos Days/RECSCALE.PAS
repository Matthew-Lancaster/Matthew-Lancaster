program recscale;


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
 min:= round(0.5*40);
 mean:= round(1.0*40);
 max:= round(5.0*40);   { xysamples for plotter are 225 max }
 end;   { of procedure initialise }









procedure makegraph;
var
xsize, ysize    : word;
xscale, yscale  : real;
xoff, yoff      : word;
x1, y1, x2, y2  : word;
x, y            : real;
s               : string[3];
 begin
 ysize:= 200;
 xsize:= 200*5 div 4;
 xscale:= xsize/(max-min-1);
 xscale:= xscale*(max-min)/max;
 yscale:= ysize/100;
 xoff:= (screenx - xsize) div 2;
 yoff:= (screeny - ysize) div 2;
 displaybox(black, xoff, yoff, xoff+xsize, yoff+ysize, 5, 4);
 disptext('THE RANTZEN SCALE', xoff+round(xsize/2), yoff-40, black, 1);
 x:= 0;
 repeat
  str(x:1:0, s);
  disptext(s, xoff+round(x*xsize/5)+2, yoff+ysize+16, black, 1);
  x:= x+1;
 until x > 5;
 disptext('paint area', xoff+round(xsize/2), yoff+ysize+40, black, 1);
 disptext('quips', xoff+round(xsize/2), yoff+ysize+52, black, 1);
 y:= 0;
 repeat
  str(y:1:0, s);
  disptext(s, xoff-16, yoff+ysize-round(y*ysize/100), black, 1);
  y:= y+25;
 until y > 100;
 disptext('rec', xoff-56, yoff+round(ysize/2), black, 1);
 disptext('%', xoff-56, yoff+16+round(ysize/2), black, 1);
 G[min]:= 0;
 setlinestyle(0, 0, 3);
 setviewport(xoff, yoff, xoff+xsize, yoff+ysize, clipon);
 for n:= min to max do
  begin
  if (n< mean) then
   G[n-min]:= (n-min)*25/(mean-min)
   else
   G[n-min]:= 25 + ln(n/mean)*75/ln(5);
   end;
 for n:= min to max-1 do
  begin
  x1:= round(n*xscale);
  y1:= round((100 - G[n-min])*yscale);
  x2:= round((n+1)*xscale);
  y2:= round((100 - G[n-min+1])*yscale);
  line(x1, y1, x2, y2);
  end;
 setlinestyle(0, 0, 1);
 end;   { of procedure makegraph }









procedure labelscale;
var
xoff, yoff      : word;
Lx, Ly          : real;
s               : string[3];
 begin
 resetplotterwindow;
 Ly:= plotarea*0.25/100;
 Lx:= Ly/1.25;
 write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2, 'LO 15');

 xoff:= midx;
 yoff:= y1 + ysize + round(ysize/8);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'THE RANTZEN SCALE', #3);

 Ly:= plotarea*0.2/100;
 Lx:= Ly/1.25;
 write(Lst, 'SI', Lx:4:2, ', ', Ly:4:2);

 yoff:= y1 - round(ysize/8);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'paint area', #3);

 yoff:= yoff - plotunits(Ly*20);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'quips', #3);

 xoff:= x1- round(xsize/5);
 yoff:= midy;
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LB', 'rec', #3);
 yoff:= yoff - plotunits(Ly*20);
 write(Lst, 'PA', xoff, ' ', yoff);
 write(Lst, 'LO 15', 'LB', '%', #3);

 xoff:= x1;
 yoff:= y1 - plotunits(Lx*10);
 for n:= 1 to 6 do
  begin
  write(Lst, 'PA', xoff, ' ', yoff);
  str(n-1, s);
  write(Lst, 'LO 16', 'LB', s, #3);
  xoff:= x1 + round(xsize*n/5);
  end;
 xoff:= x1 - plotunits(Lx*10);
 yoff:= y1;
 for n:= 1 to 5 do
  begin
  write(Lst, 'PA', xoff, ' ', yoff);
  str((n-1)*25, s);
  write(Lst, 'LO 18', 'LB', s, #3);
  yoff:= y1 + round(ysize*n/4);
  end;
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
  yscale:= ysize/100;
  L:= max - min;
  for n:= 0 to L do
   begin
   xy[n,0]:= x1 + round(min*xscale) + round(n*xscale);
   xy[n,1]:= y1 + round(G[n]*yscale);
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
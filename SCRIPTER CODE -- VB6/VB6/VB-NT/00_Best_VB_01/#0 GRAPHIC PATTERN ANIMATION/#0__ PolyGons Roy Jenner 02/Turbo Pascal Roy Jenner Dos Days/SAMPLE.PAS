program SAMPLE;

uses  CRT,  GRAPH,  KEY,  TEXTS,  ICONFILE,  TABLE,  DISPLAY,  PLOT;









var


icon            : icontype;

dispsize        : word;

papersize       : byte;

dispx, dispy    : word;

disprec         : boolean;

dotsize         : word;

dotmem          : pointer;

gotmem          : boolean;

iconsize        : real;

midx, midy      : word;

plotx, ploty    : word;

plotsize        : word;

patarea         : real;

check           : boolean;

xmin, xmax      : byte;

ymin, ymax      : byte;

xrange          : byte;

yrange          : byte;





const

up              = 0;

down            = 1;








procedure drawhelp;
var
L, lines : byte;
x, y     : word;
 begin
 L:= 19;
 lines:= 12;
 x:= 8;
 y:= 36;
 openhelpwindow(white, L, lines);
 outtextxy(x, y, 'keyA :  Area');
 inc(y,12);
 outtextxy(x, y, 'keyC :  Colours');
 inc(y,12);
 outtextxy(x, y, 'KeyE :  Edit');
 inc(y,12);
 outtextxy(x, y, 'KeyN :  Normalise');
 inc(y,12);
 outtextxy(x, y, 'keyR :  Reverse');
 inc(y,12);
 outtextxy(x, y, 'AltF :  File');
 inc(y,12);
 outtextxy(x, y, 'AltR :  Rotate');
 inc(y,12);
 outtextxy(x, y, 'AltP :  Plot');
 inc(y,12);
 outtextxy(x, y, 'AltS :  Samples');
 repeat
  presskey;
 until esc;
 closehelpwindow;
 resetkeys;
 end;   { of procedure drawhelp }








procedure edithelp;
var
L, lines : byte;
x, y     : word;
 begin
 L:= 37;
 lines:= 6;
 x:= 8;
 y:= 36;
 openhelpwindow(white, L, lines);
 outtextxy(x, y, 'Del                     :  delete dot');
 inc(y,12);
 outtextxy(x, y, 'Plus/Minus              :  move dot');
 inc(y,12);
 outtextxy(x, y, 'CurUp/CurDn/CurLt/CurRt :  edit dot');
 repeat
  presskey;
 until esc;
 closehelpwindow;
 resetkeys;
 end;   { of procedure edithelp }








procedure drawdot(xsample, ysample : byte);
var
x, y            : word;
xoff, yoff      : word;
xscale, yscale  : real;
 begin
 if not gotmem then
  begin
  getdispscales(PA, dispsize, icon.ratio, xscale, yscale, xoff, yoff);
  x:= round(xsample*xscale);
  y:= round((91 - ysample)*yscale);
  x:= x + xoff;
  y:= y + yoff;
  dotsize:= imagesize(0, 0, 5, 5);
  getmem(dotmem, dotsize);
  gotmem:= true;
  if (x> 2) and (y>2) then  getimage(x-2, y-2, x+2, y+2, dotmem^);
  setcolor(black);
  setfillstyle(1,black);
  pieslice(x, y, 0, 360, 2);
  gotmem:= true;
  end;
 end;   { of procedure drawdot }








procedure erasedot(xsample, ysample : byte);
var
x, y            : word;
xoff, yoff      : word;
xscale, yscale  : real;
 begin
 if gotmem then
  begin
  getdispscales(PA, dispsize, icon.ratio, xscale, yscale, xoff, yoff);
  x:= round(xsample*xscale);
  y:= round((91 - ysample)*yscale);
  x:= x + xoff;
  y:= y + yoff;
  if (x> 2) and (y>2) then  putimage(x-2, y-2, dotmem^, normalput);
  freemem(dotmem, dotsize);
  gotmem:= false;
  end;
 end;   { of procedure erasedot }








procedure getnextsamples(dir : byte; var x, y : byte);
var
lastcursor : word;
 begin
 lastcursor:= cursor;
 repeat
  if (dir= down)  then if (cursor< xysamples) then inc(cursor);
  if (dir= up)    then if (cursor> 1) then dec(cursor);
  x:= icon.samples[cursor,0];
  y:= icon.samples[cursor,1];
 until ((x<> patbreak) and (y<> patbreak)) or
       (cursor= 1) or (cursor= xysamples);
 if (x= patbreak) or (y= patbreak) then
  begin
  cursor:= lastcursor;
  x:= icon.samples[cursor,0];
  y:= icon.samples[cursor,1];
  end;
 end;   { of procedure getnextsamples }








procedure getfirstsamples(var x, y : byte);
 begin
 x:= icon.samples[cursor,0];
 y:= icon.samples[cursor,1];
 if (x= patbreak) or (y= patbreak) then
  begin
  getnextsamples(down, x, y);
  resetkeys;
  end;
 if (x= patbreak) or (y= patbreak) then
  begin
  getnextsamples(up, x, y);
  resetkeys;
  end;
 if (x= patbreak) or (y= patbreak) then esc:= true;
 end;   { of procedure getfirstsamples }








procedure drawgradients(x, y, color : byte);
var
x1, x2, y1, y2  : word;
xoff, yoff      : word;
xscale, yscale  : real;
 begin
 getdispscales(PA, dispsize, icon.ratio, xscale, yscale, xoff, yoff);
 setcolor(color);
 if (cursor> 1) then
  begin
  x1:= icon.samples[cursor-1, 0];
  y1:= icon.samples[cursor-1, 1];
  if (x1<>patbreak) and (y1<> patbreak) then
   begin
   x1:= round(x1*xscale) + xoff;
   y1:= round((91 - y1)*yscale) + yoff;
   x2:= round(x*xscale) + xoff;
   y2:= round((91 - y)*yscale) + yoff;
   line(x1, y1, x2, y2);
   end;
  end;
 if (cursor< xysamples) then
  begin
  x2:= icon.samples[cursor+1, 0];
  y2:= icon.samples[cursor+1, 1];
  if (x2<>patbreak) and (y2<> patbreak) then
   begin
   x1:= round(x*xscale) + xoff;
   y1:= round((91 - y)*yscale) + yoff;
   x2:= round(x2*xscale) + xoff;
   y2:= round((91 - y2)*yscale) + yoff;
   line(x1, y1, x2, y2);
   end;
  end;
 end;   { of procedure drawgradients }








procedure changesamples(var x, y : byte);
var
lastx, lasty : byte;
 begin
 lastx:= x;
 lasty:= y;
 drawgradients(x, y, white);
 if CurLt and (x> 0)  then dec(x);
 if CurRt and (x< 91) then inc(x);
 if CurDn and (y> 0)  then dec(y);
 if CurUp and (y< 91) then inc(y);
 icon.samples[cursor,0]:= x;
 icon.samples[cursor,1]:= y;
 drawgradients(x, y, black);
 if (lastx<> x) then iconedit:= true;
 if (lasty<> y) then iconedit:= true;
 end;   { of procedure changesamples }







procedure editsample;
var
x, y  : byte;
lastcursor : byte;
 begin
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 gotmem:= false;
 getfirstsamples(x, y);
 if not esc then drawdot(x, y);
 repeat
  if F1 then edithelp;
  if Del then
   begin
   erasedot(x, y);
   icon.samples[cursor,0]:= patbreak;
   icon.samples[cursor,1]:= patbreak;
   drawgradients(x, y, white);
   calcparts(icon);
   lastcursor:= cursor;
   getnextsamples(down, x, y);
   if (cursor= lastcursor) then getnextsamples(up, x, y);
   if (cursor= lastcursor) then getfirstsamples(x, y);
   drawdot(x, y);
   end;
  if Plus or Minus then
   begin
   erasedot(x, y);
   if Plus  then getnextsamples(down, x, y);
   if Minus then getnextsamples(up, x, y);
   drawdot(x, y);
   end;
  if CurLt or CurRt or CurUp or CurDn or Del then
   begin
   erasedot(x, y);
   changesamples(x, y);
   drawdot(x, y);
   end;
  if keyD then
   begin
   erasedot(x, y);
   paintcanvas(0, 0, dispsize, dispsize, white);
   displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
   drawdot(x, y);
   end;
  if not esc then presskey;
 until AltS or AltP or AltF or esc;
if esc then resetkeys;
erasedot(x, y);
end;   { of procedure editsample }








procedure getdrawpositions;
 begin
 if (icon.ratio>= 1) then
  PA:= 0.98/icon.ratio
  else
  PA:= 0.98*icon.ratio;
 dispsize:= 450;
 disprec:= false;
 dispx:= (screenx div 2) - (dispsize div 2);
 dispy:= (screeny div 2) - (dispsize div 2);
 end;   { of procedure getdrawpositions }








procedure getplotpositions;
 begin
 selectplotsize(icon.ratio, papersize, iconsize);
 papermidpoints(direction, papersize, midx, midy);
 plotsize:= plotunits(iconsize);
 plotx:= midx - (plotsize div 2);
 ploty:= midy - (plotsize div 2);
 end;   { of procedure getplotpostions }








procedure plotdrawing;
 begin
 getplotpositions;
 if CR then loadpaper(papersize);
 if CR then
  begin
  openplot(direction);
  ploticon(plotx, ploty, plotsize, 0, PA, icon.ratio, icon);
  closeplot;
  end
  else resetkeys;
 end;   { of procedure plotdrawing }








procedure reversepattern;
var
part, i : word;
 begin
 displayicon(white, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 for part:= 1 to icon.parts do
  for i:= icon.start[part] to icon.stop[part] do
   icon.samples[i,0]:= 91 - icon.samples[i,0];
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 end;   { of procedure reversepattern }









procedure rotatepattern;
var
part, i, x, y : word;
 begin
 displayicon(white, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 for part:= 1 to icon.parts do
  for i:= icon.start[part] to icon.stop[part] do
   begin
   x:= icon.samples[i,0];
   y:= 91 - icon.samples[i,1];
   icon.samples[i,0]:= y;
   icon.samples[i,1]:= x;
   end;
 icon.ratio:= 1/icon.ratio;
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 end;   { of procedure rotatepattern }









procedure measureicon;
var
part, i  : byte;
x, y     : byte;
 begin
 xmin:= icon.samples[icon.start[1],0];
 xmax:= xmin;
 ymin:= icon.samples[icon.start[1],1];
 ymax:= ymin;
 for part:= 1 to icon.parts do
  for i:= icon.start[part] to icon.stop[part] do
   begin
   x:= icon.samples[i,0];
   y:= icon.samples[i,1];
   if x< xmin then xmin:= x;
   if x> xmax then xmax:= x;
   if y< ymin then ymin:= y;
   if y> ymax then ymax:= y;
   end;
 xrange:= xmax - xmin;
 yrange:= ymax - ymin;
 end;   { of procedure measureicon }









procedure normaliseicon;
var
part, i : byte;
 begin
 displayicon(white, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 measureicon;
 for part:= 1 to icon.parts do
  for i:= icon.start[part] to icon.stop[part] do
   begin
   icon.samples[i,0]:= round( (icon.samples[i,0] - xmin)*91/xrange );
   icon.samples[i,1]:= round( (icon.samples[i,1] - ymin)*91/yrange );
   end;
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 end;   { of procedure normaliseicon }









procedure displaymessage;
 begin
 opentextwindow(white, 32, 3);
 outtextxy(8, 16, 'pattern not 91*91');
 outtextxy(8, 32, 're-edit or use keyN to normalise');
 repeat
  presskey;
 until CR or esc;
 closetextwindow;
 resetkeys;
 end;   { of procedure displaymessage }








procedure checkiconrange;
 begin
 check:= false;
 measureicon;
 if (xrange= 91) and (yrange= 91) then check:= true;
 if (xrange<> 91) or (yrange<> 91) then
  displaymessage;
 end;   { of procedure checkiconrange }









procedure displayarea(col : byte);
var
s  : string[14];
 begin
 str(icon.area:6:4, s);
 s:= 'area ' + s;
 setviewport(0, 0, 639, 479, clipon);
 disptext(s, 100, 472, col, 1);
 end;   { of procedure displayarea }








procedure measurearea;
var
xoff, yoff, n, m, p      : word;
part, i, L, scale        : word;
maxarea, count           : longint;
loop                     : boolean;
pat                      : array[1..480] of pointtype;
 begin
 displayicon(white, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 setviewport(0, 0, 639, 479, clipon);
 scale:= 5;
 xoff:= (639-(scale*91)) div 2;
 yoff:= ((479-(scale*91)) div 2) - 6;
 checkiconrange;
 if (icon.parts> 0) and check then
  begin
  for part:= 1 to icon.parts do
   begin
   L:= icon.stop[part] - icon.start[part] + 1;
   for i:= icon.start[part] to icon.stop[part] do
    begin
    pat[i-icon.start[part] +1].x:= (icon.samples[i,0]*scale) + xoff;
    pat[i-icon.start[part] +1].y:= ((91- icon.samples[i,1])*scale) + yoff;
    end;
   setfillstyle(solidfill,black);
   setcolor(black);
   loop:= false;
   if (pat[1].x = pat[L].x) and (pat[1].y = pat[L].y) then
    begin
    loop:= true;
    drawpoly(L, pat);
    fillpoly(L, pat);
    end;
   end;
  count:= 0;
  if loop then
   begin
   for n:= xoff to xoff + (scale*91) do
    for m:= yoff to yoff + (scale*91) do
     begin
     p:= getpixel(n,m);
     if (p= black) then
      begin
      inc(count);
      putpixel(n,m,white);
      end;
     end;
   maxarea:= sqr((scale*91) +1);
   icon.area:= count/maxarea;
   end;
  end;
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 displayarea(white);
 if check and loop then
  repeat
   presskey;
  until AltF or AltS or CR or esc;
 displayarea(black);
 if esc then resetkeys;
 end;   { of procedure measurearea }









procedure displaycolours(col: word);
var
s : string[14];
 begin
 str(icon.colours:2, s);
 s:= 'colours ' + s;
 setviewport(0, 0, 639, 479, clipon);
 disptext(s, 300, 472, col, 1);
 end;   { of procedure displaycolours }










procedure changecolours;
var
s  : string[14];
 begin
 displaycolours(white);
 repeat
  if CurUp or CurDn then
   begin
   if CurUp and (icon.colours> 1) then dec(icon.colours);
   if CurDn and (icon.colours< 6) then inc(icon.colours);
   displaycolours(white);
   end;
  presskey;
 until AltF or AltS or CR or esc;
 displaycolours(black);
 if esc then resetkeys;
 end;   { of procedure changeccolours }








procedure editicon;
 begin
 screencol(white);
 getdrawpositions;
 displayicon(black, dispx, dispy, dispsize, PA, disprec, icon.ratio, icon);
 displayarea(black);
 displaycolours(black);
 repeat
  if keyA then measurearea;
  if keyC then changecolours;
  if keyE then editsample;
  if keyN then normaliseicon;
  if keyR then reversepattern;
  if AltR then rotatepattern;
  if F1 then drawhelp;
  if AltF then
   begin
   editiconfile(icon);
   if iconread then
    begin
    screencol(white);
    getdrawpositions;
    displayicon(black, dispx, dispy, dispsize, PA, disprec,
                icon.ratio, icon);
    displayarea(black);
    displaycolours(black);
    cursor:= 1;
    end;
   end;
  if AltP then plotdrawing;
  if esc and iconedit then iconsavecheck(icon);
  if (not AltS) and (not esc) then presskey;
 until AltS or esc;
 end;   { of procedure editicon }








begin




plotarea:= 100;
direction:= horizontal;
reseticon(icon);
iconsize:= 100;   { mm }
papersize:= A4;
plotquality:= highquality;
patarea:= 0;
editiconfile(icon);
editicon;
repeat
 if AltD then editicon;
 if AltS then edittable(icon);
until esc;


closegraph;


end.   { of program SAMPLE }
unit TEXTS;

interface

uses   GRAPH,  KEY;








procedure opengraph;
procedure screencol(color : byte);
procedure paintcanvas(x1, y1, x2, y2  : real; color  : byte);
procedure disptext(text : string; xpos, ypos, textcol, size: word);
procedure erasetext(text : string; xpos, ypos, bkcol, size : word);
procedure openhelpwindow(color, chars, lines : byte);
procedure closehelpwindow;
procedure opentextwindow(color, chars, lines : byte);
procedure closetextwindow;
procedure closeplotwindow;
procedure openplotwindow(color, chars, lines : byte);
procedure readstring(x, y, color, maxL : word; cursor : boolean; var txt);
procedure erasestring(color : byte);
procedure erasetextcursor(backcol : byte);










var

bkgcol       : byte;

canvascol    : byte;

screenx      : word;

screeny      : word;







implementation



var

helpsize        : word;

helppoint       : pointer;

helpx1, helpx2  : word;

helpy1, helpy2  : word;

gothelpmem      : boolean;

hvp             : viewporttype;

textsize        : word;

textpoint       : pointer;

textx1, textx2  : word;

texty1, texty2  : word;

gottextmem      : boolean;

tvp             : viewporttype;

plotsize        : word;

plotpoint       : pointer;

plotx1, plotx2  : word;

ploty1, ploty2  : word;

gotplotmem      : boolean;

pvp             : viewporttype;

xcur, ycur      : word;

tx1, tx2        : word;

ty1, ty2        : word;












procedure opengraph;
var
gd, gm  : integer;
 begin
 gd:= detect;
 detectgraph(gd,gm);
 initgraph(gd,gm,'a:');
 cleardevice;
 bkgcol:= black;
 screenx:= getmaxx;
 screeny:= getmaxy;
 end;   { of procedure opengraph }







procedure screencol(color : byte);
 begin
 setviewport(0,0,639,479,clipon);
 if color= black then
 cleardevice
 else
  begin
  setfillstyle(1,1);
  bar(0,0,639,479);
  end;
 bkgcol:= color;
 end;   { of procedure screencol }








procedure paintcanvas(x1, y1, x2, y2  : real; color  : byte);
 begin
 if color= white then setfillstyle(1,1)
  else setfillstyle(0,0);
  bar(round(x1), round(y1), round(x2), round(y2));
  canvascol:= color;
  end;   { of procedure paintcanvas }








procedure disptext(text: string; xpos, ypos, textcol, size : word);
var
w, h, x, y, xs, ys, xf, yf  : word;
 begin
 if text[0]> #0 then
  begin
  settextstyle(0,0,size);
  w:= textwidth(text);
  h:= textheight(text);
  x:= xpos- ( w div 2 );
  y:= ypos - ( h div 2 );
  xs:= x - ( w div 10 );
  ys:= y - ( h div 5 );
  xf:= x + w + ( w div 10);
  yf:= y + h + ( h div 5);
  if textcol= white then setfillstyle(0,0);
  if textcol= black then setfillstyle(1,1);
  bar(xs,ys,xf,yf);
  setcolor(textcol);
  outtextxy(x,y,text);
  end;
 end;   { of disptext }









procedure erasetext(text : string; xpos, ypos, bkcol, size : word);
var
w, h, x, y, xs, ys, xf, yf  : word;
 begin
  settextstyle(0,0,size);
  w:= textwidth(text);
  h:= textheight(text);
  x:= xpos - ( w div 2 );
  y:= ypos - ( h div 2 );
  xs:= x - ( w div 10 );
  ys:= y - ( h div 5);
  xf:= x + w + ( w div 10 );
  yf:= y + h + ( h div 5 );
  if bkcol= black then setfillstyle(0,0);
  if bkcol= white then setfillstyle(1,1);
  bar(xs,ys,xf,yf);
 end;   { of procedure erasetext }








procedure closehelpwindow;
 begin
 if gothelpmem then
  begin
  setviewport(0, 0, 639, 479, clipon);
  putimage(helpx1, helpy1, helppoint^, normalput);
  freemem(helppoint, helpsize);
  gothelpmem:= false;
  setviewport(hvp.x1, hvp.y1, hvp.x2, hvp.y2, hvp.clip);
  end;
 end;   { of procedure closehelpwindow }









procedure openhelpwindow(color, chars, lines : byte);
var
xsize    : word;
ysize    : word;
 begin
 if gothelpmem then closehelpwindow;
 getviewsettings(hvp);
 setviewport(0, 0, screenx, screeny, clipon);
 xsize:= 16 + (chars*8);
 ysize:= 16 + (lines*12);
 helpx1:= screenx - (screenx div 10) - xsize;
 helpy1:= 230 - (ysize div 2);
 helpx2:= helpx1 + xsize;
 helpy2:= helpy1 + ysize;
 helpsize:= imagesize(helpx1, helpy1, helpx2, helpy2);
 getmem(helppoint, helpsize);
 gothelpmem:= true;
 getimage(helpx1, helpy1, helpx2, helpy2, helppoint^);
 setviewport(helpx1, helpy1, helpx2, helpy2, clipon);
 if color= black then
  begin
  setcolor(white);
  setfillstyle(0,1);
  end
  else
  begin
  setcolor(black);
  setfillstyle(1,1);
  end;
 bar(0, 0, helpx2-helpx1, helpy2-helpy1);
 if color= bkgcol then
  rectangle(0, 0, helpx2-helpx1, helpy2-helpy1);
 outtextxy(8, 12, 'HELP !');
 end;   { of procedure openhelpwindow }









procedure closetextwindow;
 begin
 if gottextmem then
  begin
  setviewport(0, 0, 639, 479, clipon);
  putimage(textx1, texty1, textpoint^, normalput);
  freemem(textpoint, textsize);
  gottextmem:= false;
  setviewport(tvp.x1, tvp.y1, tvp.x2, tvp.y2, tvp.clip);
  end;
 end;   { of procedure closetextwindow }









procedure opentextwindow(color, chars, lines : byte);
var
xsize    : word;
ysize    : word;
 begin
 if gottextmem then closetextwindow;
 getviewsettings(tvp);
 setviewport(0, 0, 639, 479, clipon);
 xsize:= 16 + (chars*8);
 ysize:= 16 + (lines*12);
 textx1:= 320 - (xsize div 2);
 texty1:= 230 - (ysize div 2);
 textx2:= textx1 + xsize;
 texty2:= texty1 + ysize;
 textsize:= imagesize(textx1, texty1, textx2, texty2);
 getmem(textpoint, textsize);
 gottextmem:= true;
 getimage(textx1, texty1, textx2, texty2, textpoint^);
 setviewport(textx1, texty1, textx2, texty2, clipon);
 if color= black then
  begin
  setcolor(white);
  setfillstyle(0,1);
  end
  else
  begin
  setcolor(black);
  setfillstyle(1,1);
  end;
 bar(0, 0, textx2-textx1, texty2-texty1);
 if color= bkgcol then
  rectangle(0, 0, textx2-textx1, texty2-texty1);
 end;   { of procedure opentextwindow }









procedure closeplotwindow;
 begin
 if gotplotmem then
  begin
  setviewport(0, 0, 639, 479, clipon);
  putimage(plotx1, ploty1, plotpoint^, normalput);
  freemem(plotpoint, plotsize);
  gotplotmem:= false;
  setviewport(pvp.x1, pvp.y1, pvp.x2, pvp.y2, pvp.clip);
  end;
 end;   { of procedure closeplotwindow }









procedure openplotwindow(color, chars, lines : byte);
var
xsize    : word;
ysize    : word;
 begin
 if gotplotmem then closeplotwindow;
 getviewsettings(pvp);
 setviewport(0, 0, 639, 479, clipon);
 xsize:= 16 + (chars*8);
 ysize:= 16 + (lines*12);
 plotx1:= 50;
 ploty1:= (screeny div 2) - (ysize div 2);
 plotx2:= plotx1 + xsize;
 ploty2:= ploty1 + ysize;
 plotsize:= imagesize(plotx1, ploty1, plotx2, ploty2);
 getmem(plotpoint, plotsize);
 gotplotmem:= true;
 getimage(plotx1, ploty1, plotx2, ploty2, plotpoint^);
 setviewport(plotx1, ploty1, plotx2, ploty2, clipon);
 if color= black then
  begin
  setcolor(white);
  setfillstyle(0,1);
  end
  else
  begin
  setcolor(black);
  setfillstyle(1,1);
  end;
 bar(0, 0, plotx2-plotx1, ploty2-ploty1);
 if color= bkgcol then
  rectangle(0, 0, plotx2-plotx1, ploty2-ploty1);
 end;   { of procedure openplotwindow }









procedure erasetextcursor(backcol : byte);
 begin
 if backcol= white then
  setfillstyle(1,1)
  else
  setfillstyle(0,0);
 bar(xcur, ycur, xcur+6, ycur+1);
 end;   { of procedure erasetextcursor }










procedure displaytextcursor(curcol : word);
 begin
 if curcol= white then
  setfillstyle(1,1)
  else
  setfillstyle(0,0);
 bar(xcur, ycur, xcur+6, ycur+1);
 end;   { of procedure displaytextcursor }










procedure erasestring(color : byte);
 begin
 if color= black then
 setfillstyle(0,0)
 else
 setfillstyle(1,1);
 bar(tx1, ty1, tx2, ty2);
 end;   { of procedure erasestring }








procedure readstring(x, y, color, maxL : word; cursor : boolean; var txt);
var
cur    : byte;
type
t    = string;
 begin
 setcolor(color);
 if color= black then
  setfillstyle(1,1)
  else
  setfillstyle(0,0);
 cur:= ord(t(txt)[0]);
 if cur> maxL then cur:= maxL;
 ycur:= y + 9;
 if not funkey and (cur < maxL) then
   begin
   if (ch>= #32) and (ch<= #126) then
    begin
    t(txt):= t(txt) + ch;
    inc(cur);
    end;
   end;
 if Del then
  begin
  ord(t(txt)[0]):= 0;
  cur:= 0;
  end;
 if BkSp and (cur> 0) then
  begin
  dec(ord(t(txt)[0]));
  dec(cur);
  end;
 if cursor then
  begin
  tx1:= x-4;
  ty1:= y-3;
  tx2:= x + (maxL*8) + 6;
  ty2:= y + 12;
  end
  else
  begin
  tx1:= x-3;
  ty1:= y-3;
  tx2:= x + (maxL*8) + 1;
  ty2:= y + 11;
  end;
 bar(tx1, ty1, tx2, ty2);
 outtextxy(x, y, t(txt));
 xcur:= x + (cur*8) -1;
 if cursor then displaytextcursor(color);
 end;   { of procedure readstring }










begin



gothelpmem:= false;
gottextmem:= false;
gotplotmem:= false;


end.   { of unit TEXTS }
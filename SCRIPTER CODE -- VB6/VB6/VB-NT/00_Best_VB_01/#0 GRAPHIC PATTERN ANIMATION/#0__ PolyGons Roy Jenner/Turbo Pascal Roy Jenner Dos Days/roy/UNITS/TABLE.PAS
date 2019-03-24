unit  TABLE;

interface

uses  DOS,  GRAPH,  KEY,  TEXTS,  ICONFILE;







procedure calcparts(var icon : icontype);
procedure edittable(var icon : icontype);







var

cursor   : word;






implementation


var

xyrow    : byte;

tablex   : word;

tabley   : word;

space    : word;

w, h     : byte;

namex    : word;

namey    : word;

ratiox   : word;

ratioy   : word;









procedure samplehelp;
var
L, lines : byte;
x, y     : word;
 begin
 L:= 15;
 lines:= 7;
 x:= 8;
 y:= 36;
 openhelpwindow(white, L, lines);
 outtextxy(x, y, 'AltD :  DRAWING');
 inc(y,12);
 outtextxy(x, y, 'AltF :  FILE');
 inc(y,12);
 outtextxy(x, y, 'AltN :  NAME');
 inc(y,12);
 outtextxy(x, y, 'AltR :  RATIO');
 repeat
  presskey;
 until esc;
 closehelpwindow;
 resetkeys;
 end;   { of procedure samplehelp }








procedure partsmessage;
 begin
 opentextwindow(white, 14, 2);
 setcolor(black);
 outtextxy(8,  8,   'TOO MANY PARTS');
 outtextxy(8,  26,  'press esc');
 repeat
  presskey;
 until esc;
 closetextwindow;
 resetkeys;
 end;   { of procedure partsmessage }







procedure calcparts(var icon : icontype);
var
i, j        : byte;
part        : byte;
x, y        : byte;
x1, y1      : byte;
 begin
  icon.parts:= 0;
  part:= 0;
  i:= 0;
  repeat
   inc(i);
   x:=  icon.samples[i,0];
   y:=  icon.samples[i,1];
   x1:= icon.samples[i+1,0];
   y1:= icon.samples[i+1,1];
   if (x<> patbreak) and (y<> patbreak) and
      (x1<> patbreak) and (y1<> patbreak) then
    begin
    inc(part);
    if part<= maxparts then
     inc(icon.parts)
     else
     partsmessage;
    icon.start[part]:= i;
    j:= i+1;
    repeat
     inc(j);
     x:= icon.samples[j, 0];
     y:= icon.samples[j, 1];
    until (x= patbreak) or (y= patbreak) or (j > xysamples-1);
    icon.stop[part]:= j-1;
    i:= j;
    end;
   until i> (xysamples-2);
 end;   { of procedure calcparts }








procedure locatecursor(cursor : word; var row, loc : word);
 begin
 loc:= cursor;
 if (loc > 114) then loc:= loc - 114;
 row:= 1;
 repeat
  if (loc > 19) then
   begin
   inc(row);
   dec(loc, 19);
   end;
 until loc <= 19;
 end;   { of procedure locatecursor }







procedure printname(icon  : icontype);
 begin
 namex:= tablex + (8*8);
 namey:= tabley + round(6.8*space);
 setfillstyle(1,1);
 bar(namex, namey, namex+(12*8), namey+7);
 outtextxy(namex, namey, icon.name);
 end;   { of procedure printname }








procedure printratio(icon : icontype);
var
s     : string[4];
 begin
 ratiox:= tablex + (20*w) - (12*8);
 ratioy:= tabley + round(6.8*space);
 str(icon.ratio:2:2, s);
 setfillstyle(1,1);
 bar(ratiox, ratioy, ratiox+(4*8), ratioy+7);
 outtextxy(ratiox, ratioy, s);
 end;   { of procedure printratio }








procedure printxy(cursor : word; icon : icontype);
var
x, y  : byte;
row   : word;
loc   : word;
xoff  : word;
yoff  : word;
xstr  : string[2];
ystr  : string[2];
 begin
 locatecursor(cursor, row, loc);
 xoff:= tablex + 4 + (loc*w);
 yoff:= tabley + (row*space);
 setfillstyle(1,1);
 setcolor(black);
 bar(xoff, yoff-28, xoff+14, yoff-28+6);
 bar(xoff, yoff-12, xoff+14, yoff-12+6);
 x:= icon.samples[cursor,0];
 y:= icon.samples[cursor,1];
 if (x<> patbreak) then
  begin
  str(x, xstr);
  outtextxy(xoff, yoff-28, xstr);
  end;
 if (y<> patbreak) then
  begin
  str(y, ystr);
  outtextxy(xoff, yoff-12, ystr);
  end;
 end;   { of procedure printxy }








procedure printtable(icon : icontype);
var
loc   : word;
start : word;
stop  : word;
 begin
 setfillstyle(1,1);
 setcolor(black);
 bar(600, 450, 608, 457);
 if (cursor<= 114) then   { page 1 }
  begin
  start:= 1;
  stop:= 114;
  outtextxy(600, 450, '1/2');
  end;
 if (cursor> 114) then    { page 2 }
  begin
  start:= 115;
  stop:= 228;
  outtextxy(600, 450, '2/2');
  end;
 for loc:= start to stop do
  printxy(loc, icon);
 printname(icon);
 printratio(icon);
 end;   { of procedure printtable }








procedure drawxyrow(row : word);
var
i    : byte;
yoff : word;
 begin
 yoff:= tabley + (row*space);
 for i:= 0 to 2 do
  line(tablex, yoff-(i*h), tablex + (20*w), yoff-(i*h));
 for i:= 0 to 20 do
  line(tablex+(w*i), yoff, tablex+(w*i), yoff-(2*h));
 outtextxy(tablex+8, yoff-28, 'x');
 outtextxy(tablex+8, yoff-12, 'y');
 end;   { of procedure drawxyrow }







procedure drawtable;
var
row : byte;
 begin
 screencol(white);
 setcolor(black);
 setviewport(0, 0, screenx, screeny, clipon);
 w:= 22;
 h:= 16;
 tablex:= (screenx - (20*w)) div 2;
 tabley:= 40;
 space:= round(h*3.4);
 for row:= 1 to 6 do
  drawxyrow(row);
 end;   { of procedure drawtable }








procedure readname(var icon : icontype);
 begin
 resetkeys;
 repeat
  readstring(namex, namey, black, 12, true, icon.name);
  if F1 then samplehelp;
  presskey;
 until CR or AltR or AltF or AltD or esc;
 erasetextcursor(white);
 end;   { of procedure readname }








procedure readratio(var icon : icontype);
var
code  : integer;
s     : string[4];
 begin
 resetkeys;
 str(icon.ratio:2:2, s);
 repeat
  readstring(ratiox, ratioy, black, 4, true, s);
  val(s, icon.ratio, code);
  if F1 then samplehelp;
  presskey;
 until (CR and (code=0)) or (AltN and (code= 0)) or
       (AltF and (code= 0)) or (AltD and (code= 0)) or (esc and (code= 0));
 erasetextcursor(white);
 end;   { of procedure readratio }








procedure readnameratio(var icon : icontype);
 begin
 repeat
  if AltN then readname(icon);
  if AltR then readratio(icon);
 until CR or AltF or AltD or esc;
 end;   { of procedure readnameratio }








procedure editsample(var icon : icontype);
var
row    : word;
loc    : word;
xoff   : word;
yoff   : word;
code   : integer;
xystr  : string[2];
 begin
 locatecursor(cursor, row, loc);
 xoff:= tablex + 4 + (loc*w);
 yoff:= tabley + (row*space);
 if xyrow= 0 then yoff:= yoff - 28;
 if xyrow= 1 then yoff:= yoff - 12;
 if icon.samples[cursor,xyrow] <> patbreak then
  str(icon.samples[cursor,xyrow], xystr)
  else
  ord(xystr[0]):= 0;
 readstring(xoff, yoff, white, 2, false, xystr);
 if ord(xystr[0]) <>  0 then
  val(xystr, icon.samples[cursor,xyrow], code)
  else
  icon.samples[cursor,xyrow]:= patbreak;
 if (ch> #47) and (ch< #58) then iconedit:= true;
 end;   { of procedure editsample }








procedure editcursor(var icon : icontype);
var
row    : word;
loc    : word;
 begin
 locatecursor(cursor, row, loc);
 if PgUp and (cursor> 114) then
  begin
  dec(cursor, 114);
  printtable(icon);
  end;
 if PgDn and (cursor <= 114) then
  begin
  inc(cursor, 114);
  printtable(icon);
  end;
 if Home then
  begin
  if (cursor<= 114) then cursor:= 1;
  if (cursor> 114)  then cursor:= 115;
  xyrow:= 0;
  end;
 if En then
  begin
  if (cursor<= 114) then cursor:= 114;
  if (cursor> 114)  then cursor:= 228;
  xyrow:= 0;
  end;
 if CurRt and (cursor< xysamples) then inc(cursor);
 if CurLt and (cursor> 1) then dec(cursor);
 if CurLt and (cursor= 114) then printtable(icon);
 if CurRt and (cursor= 115) then printtable(icon);
 if CurDn and (xyrow= 1) and (row< 6) then
  begin
  xyrow:= 0;
  inc(cursor, 19);
  resetkeys;
  end;
 if CurUp and (xyrow= 0) and (row> 1) then
  begin
  xyrow:= 1;
  dec(cursor, 19);
  resetkeys;
  end;
 if CurUp and (xyrow= 1) then dec(xyrow);
 if CurDn and (xyrow= 0) then inc(xyrow);
 editsample(icon);
 calcparts(icon);
 end;   { of procedure editcursor }









procedure edittable(var icon : icontype);
 begin
 drawtable;
 printtable(icon);
 repeat
  if F1 then samplehelp;
  if AltN or AltR then readnameratio(icon);
  if AltF then
   begin
   editiconfile(icon);
   if iconread then
    begin
    xyrow:= 0;
    cursor:= 1;
    printtable(icon);
    end;
   end;
  editcursor(icon);
  if not AltD then presskey;
  erasestring(white);
  printxy(cursor, icon);
  if esc and iconedit then iconsavecheck(icon);
 until AltD or esc;
 end;   { of procedure edittable }








begin


xyrow:= 0;   { x }
cursor:= 1;



end.   { of unit  TABLE }
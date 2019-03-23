unit  ICONFILE;

interface

uses   DOS,   GRAPH,   KEY,  TEXTS;










const


maxparts  = 20;

xysamples = 228;  { 114*2 = 2 pages}

patbreak  = 255;





type



icontype    =  record

               name     : string[12];

               notes    : string[50];

               parts    : byte;

               ratio    : real;

               area     : real;

               colours  : byte;

               start    : array[1..maxparts] of word;

               stop     : array[1..maxparts] of word;

               samples  : array[1..xysamples, 0..1] of byte;

               end;









procedure editiconfile(var icon : icontype);
procedure iconsavecheck(icon : icontype);
procedure reseticon(var icon : icontype);
procedure writeicon(var icon : icontype);
procedure readicon(var icon : icontype);






var

iconread         : boolean;

iconedit         : boolean;







implementation





var


f                : text;

filename         : string[21];

writecheck       : boolean;

filesize         : word;

filex1, filex2   : word;

filey1, filey2   : word;

filepoint        : pointer;

gotfilemem       : boolean;

vp               : viewporttype;









procedure filehelp;
var
L, lines : byte;
x, y     : word;
 begin
 L:= 15;
 lines:= 5;
 x:= 8;
 y:= 36;
 openhelpwindow(white, L, lines);
 outtextxy(x, y, 'F2 :  SAVE');
 inc(y,12);
 outtextxy(x, y, 'CR :  READ');
 repeat
  presskey;
 until esc;
 closehelpwindow;
 resetkeys;
 end;   { of procedure filehelp }








procedure clearfilewindow;
 begin
 paintcanvas(1, 1, filex2-filex1-1, filey2-filey1-1, white);
 end;   { of procedure clearfilewindow }








procedure closefilewindow;
 begin
 setviewport(0, 0, screenx, screeny, clipon);
 putimage(filex1, filey1, filepoint^, normalput);
 freemem(filepoint, filesize);
 setviewport(vp.x1, vp.y1, vp.x2, vp.y2, vp.clip);
 gotfilemem:= false;
 end;   { of procedure closefilewindow }








procedure openfilewindow;
 begin
 filex1:= 70;
 filex2:= screenx - 70;
 filey1:= 40;
 filey2:= 74;
 if gotfilemem then closefilewindow;
 getviewsettings(vp);
 setviewport(0, 0, screenx, screeny, clipon);
 filesize:= imagesize(filex1, filey1, filex2, filey2);
 getmem(filepoint, filesize);
 gotfilemem:= true;
 getimage(filex1, filey1, filex2, filey2, filepoint^);
 setcolor(black);
 rectangle(filex1, filey1, filex2, filey2);
 setviewport(filex1, filey1, filex2, filey2, clipon);
 clearfilewindow;
 end;   { of procedure openfilewindow }








procedure filemessage(message  : string);
 begin
 clearfilewindow;
 disptext(message, (filex2-filex1-8) div 2, 14, black,1);
 end;   { of procedure filemessage }









procedure printname(icon : icontype);
 begin
 outtextxy(12,        6,  'NAME   :  ');
 outtextxy(12+ (9*8), 6,  icon.name);
 end;   { of procedure printname }









procedure printnotes(icon : icontype);
 begin
 outtextxy(12,        20, 'NOTES  :  ');
 outtextxy(12+ (9*8), 20,  icon.notes);
 end;   { of procedure printnotes }









procedure printfilewindow(icon : icontype);
 begin
 clearfilewindow;
 printname(icon);
 printnotes(icon);
 end;   { of procedure printfilewindow }









procedure editname(var icon : icontype);
 begin
 readstring(12+ (9*8), 6, black, 12, true, icon.name);
 end;   { of procedure editname }








procedure editnotes(var icon : icontype);
 begin
 readstring(12+ (9*8), 20, black, 50, true, icon.notes);
 end;   { of procedure editnotes }








procedure reseticon(var icon : icontype);
var
i  : word;
 begin
 ord(icon.name[0]):= 0;
 ord(icon.notes[0]):= 0;
 icon.ratio:= 1.0;
 icon.parts:= 0;
 icon.area:= 0;
 icon.colours:= 1;
 for i:= 1 to xysamples do
  begin
  icon.samples[i,0]:= patbreak;
  icon.samples[i,1]:= patbreak;
  end;
 end;   { of procedure reseticon }








procedure overwritecheck(icon : icontype);
var
dirinfo : searchrec;
 begin
 writecheck:= true;
 findfirst(filename, anyfile, dirinfo);
 if (dirinfo.name= icon.name) then
  begin
  filemessage( 'overwrite ' + icon.name + '?  Y/N');
  repeat
  presskey;
  until keyY or keyN or esc;
  if keyN or esc then writecheck:= false;
  printfilewindow(icon);
  resetkeys;
  end;
 end;   { of procedure overwritecheck }









procedure readicon(var icon : icontype);
var
part, i, io  : integer;
 begin
 iconread:= false;
 {$I-}
 filename:= 'A:/PATTERNS/' + icon.name;
 assign(f,filename);
 reset(f);
 close(f);
 {$I-}
 io:= ioresult;
 if io = 0 then
  begin
  assign(f,filename);
  reset(f);
  reseticon(icon);
  readln(f, icon.name);
  readln(f, icon.notes);
  readln(f, icon.parts);
  readln(f, icon.ratio);
  readln(f, icon.area);
  readln(f, icon.colours);
  for part:= 1 to icon.parts do
   begin
   readln(f, icon.start[part]);
   readln(f, icon.stop[part]);
   end;
  for part:= 1 to icon.parts do
   for i:= icon.start[part] to icon.stop[part] do
   begin
   readln(f, icon.samples[i, 0]);
   readln(f, icon.samples[i, 1]);
   end;
  close(f);
  iconread:= true;
  iconedit:= false;
  end;
 if io<> 0 then
  begin
  filemessage('cant find file  -  press any key');
  presskey;
  resetkeys;
  printfilewindow(icon);
  end;
 end;   { of procedure readicon }









procedure writeicon(var icon : icontype);
var
part, i, io  : integer;
 begin
 filename:= 'A:/PATTERNS/' + icon.name;
 overwritecheck(icon);
 if writecheck then
  begin
  {$I-}
  assign(f,filename);
  rewrite(f);
  close(f);
  {$I+}
  io:= ioresult;
  if io= 0 then
   begin
   assign(f,filename);
   rewrite(f);
   writeln(f, icon.name);
   writeln(f, icon.notes);
   writeln(f, icon.parts);
   writeln(f, icon.ratio:4:2);
   writeln(f, icon.area:6:4);
   writeln(f, icon.colours);
   for part:= 1 to icon.parts do
    begin
    writeln(f, icon.start[part]);
    writeln(f, icon.stop[part]);
    end;
   for part:= 1 to icon.parts do
    for i:= icon.start[part] to icon.stop[part] do
    begin
    writeln(f, icon.samples[i, 0]);
    writeln(f, icon.samples[i, 1]);
    end;
   close(f);
   iconedit:= false;
   end;
  if io<> 0 then
   begin
   filemessage('cant write to disk  -  press any key');
   presskey;
   resetkeys;
   printfilewindow(icon);
   end;
  end;
 end;   { of procedure writeicon }









procedure iconsavecheck(icon : icontype);
 begin
 openfilewindow;
 filemessage( 'save ' + icon.name + '?  Y/N');
 repeat
  presskey;
 until keyY or keyN or esc;
 if keyY then writeicon(icon);
 iconedit:= false;
 resetkeys;
 closefilewindow;
 end;   { of procedure iconsavecheck }








procedure editiconfile(var icon : icontype);
const
name  = 1;
notes = 2;
var
line  : byte;
 begin
 iconread:= false;
 openfilewindow;
 printfilewindow(icon);
 line:= name;
 repeat
  if CurDn and (line= name) then
   begin
   erasetextcursor(white);
   line:= notes;
   end;
  if CurUp and (line= notes) then
   begin
   erasetextcursor(white);
   line:= name;
   end;
  if F1 then filehelp;
  if F2 then writeicon(icon);
  if CR then readicon(icon);
  if line= name  then editname(icon);
  if line= notes then editnotes(icon);
  if not iconread then presskey;
  until iconread or esc;
 closefilewindow;
 resetkeys;
 end;   { of procedure editiconfile }









begin



gotfilemem:= false;
iconedit:= false;


end.  { of unit  ICONFILE }
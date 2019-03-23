program ladder;

uses    GRAPH,  KEY,  TEXTS,  DISPLAY,  PLOT;









var

midx, midy  : word;

size        : word;

th          : real;










procedure plotladder(x, y, H, th : word);
var
size, H2, W2, G2  : word;
t, i, p2          : integer;
finish            : boolean;
 begin
 plotquality:= highquality;
 selectpen;
 size:= (H div 8)*8;
 H2:= size div 2;
 W2:= size div 4;
 G2:= size div 8;
 H2:= H2 - round(0.5*20);   { 1/2  0.5mm pen width }
 t:= th div 2;
 p2:= plotunits(0.5) div 2;
 i := -t;
 finish:= false;
  repeat
  i:= i + p2;
  if i> (t-p2) then
   begin
   i:= t-p2;
   finish:= true;
   end;
  write(Lst, 'PA', x-W2+i, ' ',  y+G2, 'PD', x-W2+i, ' ', y-H2, 'PU');
  write(Lst, 'PA', x+W2+i, ' ',  y-G2, 'PD', x+W2+i, ' ', y+H2, 'PU');
  write(Lst, 'PA', x+W2+(t-p2), ' ',  y+G2+i, 'PD',
                   x-W2-(t-p2), ' ',  y+G2+i, 'PU');
  write(Lst, 'PA', x-W2-(t-p2), ' ',  y-G2+i, 'PD',
                   x+W2+(t-p2), ' ',  y-G2+i, 'PU');
  until finish;
 end;   { of procedure plotladder }









procedure plotladder1(x, y, H, th : word);
var
size, H2, W2, G2  : word;
gap               : real;
t, i, p2          : integer;
finish            : boolean;
 begin
 plotquality:= highquality;
 selectpen;
 size:= (H div 8)*8;
 gap:= 0.236;
 H2:= size div 2;
 W2:= round(H*(0.5+(gap/2))/2);
 G2:= round(H*gap/2);
 H2:= H2 - round(0.5*20);   { 1/2  0.5mm pen width }
 t:= th div 2;
 p2:= plotunits(0.5) div 2;
 i := -t;
 finish:= false;
  repeat
  i:= i + p2;
  if i> (t-p2) then
   begin
   i:= t-p2;
   finish:= true;
   end;
  write(Lst, 'PA', x-W2+i, ' ',  y+G2, 'PD', x-W2+i, ' ', y-H2, 'PU');
  write(Lst, 'PA', x+W2+i, ' ',  y-G2, 'PD', x+W2+i, ' ', y+H2, 'PU');
  write(Lst, 'PA', x+W2+(t-p2), ' ',  y+G2+i, 'PD',
                   x-W2-(t-p2), ' ',  y+G2+i, 'PU');
  write(Lst, 'PA', x-W2-(t-p2), ' ',  y-G2+i, 'PD',
                   x+W2+(t-p2), ' ',  y-G2+i, 'PU');
  until finish;
 end;   { of procedure plotladder1 }









procedure ladderplot;
var
x, y  : word;
 begin
 loadpaper(A4);
 direction:= horizontal;
 if CR then
  begin
  papermidpoints(direction, A4, x, y);
  openplot(direction);
  plotladder1(x, y, plotunits(size), plotunits(th));
  closeplot;
  end
  else resetkeys;
 end;   { of procedure ladderplot }









procedure displayladder(x, y, size, th, colour : word);
var
H2, W2, G2 : word;
t, i       : integer;
 begin
 setviewport(0, 0, screenx, screeny, clipon);
 setcolor(colour);
 size:= (size div 8)*8;
 gap:= 0.236;
 H2:= size div 2;
 W2:= size div 4;
 G2:= size div 8;
 t:= th div 2;
 for i:= -t to t do
  begin
  line(x-W2+i,  y+H2, x-W2+i, y-G2);
  line(x+W2+i,  y-H2, x+W2+i, y+G2);
  line(x-W2-t,  y-G2+i, x+W2+t, y-G2+i);
  line(x-W2-t,  y+G2+i, x+W2+t, y+G2+i);
  end;
 end;   { of procedure displayladder }








procedure displayladder1(x, y, size, th, colour : word);
var
H2, W2, G2 : word;
gap        : real;
t, i       : integer;
 begin
 setviewport(0, 0, screenx, screeny, clipon);
 setcolor(colour);
 size:= (size div 8)*8;
 gap:= 0.236;
 H2:= size div 2;
 W2:= round(size*(0.5+(gap/2))/2);
 G2:= round(size*gap/2);
 t:= th div 2;
 for i:= -t to t do
  begin
  line(x-W2+i,  y+H2, x-W2+i, y-G2);
  line(x+W2+i,  y-H2, x+W2+i, y+G2);
  line(x-W2-t,  y-G2+i, x+W2+t, y-G2+i);
  line(x-W2-t,  y+G2+i, x+W2+t, y+G2+i);
  end;
 end;   { of procedure displayladder1 }









procedure displaysize;
var
s   : string[12];
 begin
 str(size, s);
 s:= s + ' mm';
 disptext(s, midx-50, 460, black, 1);
 str(th:5:2, s);
 s:= s + ' mm';
 disptext(s, midx+50, 460, black, 1);
 end;   { of procedure displaysize }








procedure selectladder;
var
scale       : real;
 begin
 scale:= 3.33;
 size:= 20;  { mm }
 th:= size*0.06;
 midx:= screenx div 2;
 midy:= screeny div 2;
 repeat
  displayladder1(midx, midy, round(size*scale), round(th*scale), black);
  displaysize;
  presskey;
  if Plus or Minus then
   begin
   displayladder1(midx, midy, round(size*scale), round(th*scale), white);
   repeat
    if Minus and (size> 8)   then dec(size);
    if Plus  and (size< 100) then inc(size);
    th:= size*0.06;
    displaysize;
    presskey
   until CR or esc;
   end;
  if AltP then ladderplot;
 until esc;
 end;   { of procedure selectladder }









begin



selectladder;


closegraph;


end.   { of program labell }
program paint;

uses  graph,  key,  texts,  iconfile,  display,  plot,  cel;









var

gridrange  : real;

canratio   : real;









procedure initialise;
var
off  : word;
const
A3size = 200;
A4size = 150;
 begin
 iconbox:= 34;
 columns:= 4;
 rows:= 7;
 if columns> maxcols then columns:= maxcols;  { 4 }
 if rows> maxrows then rows:= maxrows;        { 7 }
 iconxoff:= iconbox;
 iconyoff:= ((screeny - (rows*iconbox)) div 2) - 20;
 off:= (columns*iconbox) + iconxoff;
 xmid:= round(((639 - off)/2) + off);
 ymid:= 479 div 2;
 readicons;
 displayarea:= 2*rows*columns*sqr(iconbox);
 plotarea:= A3size;
 data:= false;
 maxratio:= 1.5;
 cansize:= 24;   { inches }
 end;   { of procedure initialise }








procedure makecan;
var
step   : real;
size   : real;
 begin
 size:= 0;
 step:= 0.1;
 if gridratio>= 1.0 then
  begin
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(1-gridratio)) + gridratio)*gridrange;
  size:= size - step;
  step:= 0.01;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(1-gridratio)) + gridratio)*gridrange;
  size:= size - step;
  step:= 0.001;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(1-gridratio)) + gridratio)*gridrange;
  size:= size - step;
  step:= 0.0001;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(1-gridratio)) + gridratio)*gridrange;
  end
  else
  begin
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(gridratio-1)) + 1.0)*gridrange/gridratio;
  size:= size - step;
  step:= 0.01;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(gridratio-1)) + 1.0)*gridrange/gridratio;
  size:= size - step;
  step:= 0.001;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(gridratio-1)) + 1.0)*gridrange/gridratio;
  size:= size - step;
  step:= 0.0001;
  repeat
   size:= size + step;
  until sqr(size) >= ((size*(gridratio-1)) + 1.0)*gridrange/gridratio;
  end;
 gap:= (1.0 - size)/2;
 if gridratio>= 1.0 then
  begin
  canH:= 1.0;
  gridH:= size;
  gridW:= size/gridratio;
  canW:= gridW + (2*gap);
  end
  else
  begin
  canW:= 1.0;
  gridW:= size;
  gridH:= size*gridratio;
  canH:= gridH + (2*gap);
  end;
 canratio:= canH/canW;
 end;   { of procedure makecan }










procedure makecanvas;
 begin
 gridH:= ycells*cellH;
 gridW:= xcells*cellW;
 gridrange:= 0.7416;
 gridratio:= gridH/gridW;
 makecan;
 canvasH:= round(cansize*sqrt(canratio));
 canvasW:= round(cansize/sqrt(canratio));
 end;   { of procedure makecanvas }







procedure newpainting;
 begin
 makecell;
 if (colours= 1) then cells:= 56;
 if (colours= 2) then cells:= 42;
 calcells(Down);
 makecanvas;
 makepainting;
 end;   { of procedure newpainting }









procedure calctemplate;
var
cellarea  : real;
 begin
 readcanvassizes;
 cellarea:= canH*canW*gridrange/xycells;
 tempH:= sqrt(cellarea*cellratio);
 tempW:= tempH/cellratio;
 gridarea:= tempH*tempW*xycells;
 if CR then
  begin
  if (tempH>= tempW) then
   plotsize:= tempH
   else
   plotsize:= tempW;
  plottemplate;
  end;
 if esc then resetkeys;
 end;   { of procedure calctemplate }









procedure changepicratio;
var
ratio   : real;
s, s1   : string[15];
 begin
 repeat
  if (canvasH>= canvasW) then
   ratio:= canvasH/canvasW
   else
   ratio:= canvasW/canvasH;
  s:= 'picratio : ';
  str(ratio:4:2, s1);
  s:= s + s1;
  disptext(s, 550, 40, black, 1);
  s:= 'maxratio : ';
  str(maxratio:4:2, s1);
  s:= s + s1;
  disptext(s, 550, 60, black, 1);
  presskey;
  if plus  then maxratio:= maxratio + 0.05;
  if minus then maxratio:= maxratio - 0.05;
  if CR then newpainting;
 until esc or CtrlCurLt or AltP;
 paintcanvas(450, 30, 630, 70, white);
 if esc then resetkeys;
 end;   { of procedure changepicratio }









procedure editpainting;
 begin
 editicons;
 if CR then newpainting;
 if not esc then
 repeat
  if F1 then painthelp;
  if keyD then
   begin
   data:= not data;
   if data then displaydata else erasedata;
   end;
  if keyG then
   begin
   grid:= not grid;
   if grid then
    displaybox(black, x1, y1, x2, y2, xcells, ycells)
    else
    displaybox(white, x1, y1, x2, y2, xcells, ycells);
   end;
  if keyR then reversepattern;
  if keyP then changepicratio;
  if AltP then plotpainting;
  if keyT then calctemplate;
  if CurLt or CurRt or CurUp or CurDn then
   begin
   if CurLt then
    if (xcells> 1) then xcells:= xcells - 1;
   if CurRt then
    xcells:= xcells + 1;
   if CurUp then
    if (ycells> 1) then ycells:= ycells - 1;
   if CurDn then
    ycells:= ycells + 1;
   xycells:= xcells*ycells;
   cells:= xycells;
   makecanvas;
   makepainting;
   end;
  if CtrlCurLt then
   begin
   editicons;
   if CR then newpainting;
   end;
  presskey;
 until esc;
 end;   { of procedure editpainting }









begin

initialise;

editpainting;


closegraph;

end.   { of program multiple }

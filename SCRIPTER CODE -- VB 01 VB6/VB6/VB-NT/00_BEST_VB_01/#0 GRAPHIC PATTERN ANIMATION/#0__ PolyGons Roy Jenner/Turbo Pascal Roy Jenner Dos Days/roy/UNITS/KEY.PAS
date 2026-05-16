unit KEY;

interface

uses CRT;






procedure presskey;
procedure resetkeys;






var

keyA, keyB, keyC, keyD, keyE, keyF, keyG, keyH, keyI, keyJ, keyK, keyL,
keyM, keyN, keyO, keyP, keyQ, keyR, keyS, keyT, keyU, keyV, keyW, keyX,
keyY,keyZ                                                    : boolean;

AltA, AltB, AltC, AltD, AltE, AltF, AltG, AltH, AltI, AltJ, AltK, AltL,
AltM, AltN, AltO, AltP, AltQ, AltR, AltS, AltT, AltU, AltV, AltW, AltX,
AltY, AltZ                                                   : boolean;

CtrlA, CtrlB, CtrlC, CtrlD, CtrlE, CtrlF, CtrlG, CtrlH,
CtrlI, CtrlJ, CtrlK, CtrlL, CtrlM, CtrlN, CtrlO, CtrlP,
CtrlQ, CtrlR, CtrlS, CtrlT, CtrlU, CtrlV, CtrlW, CtrlX,
CtrlY, CtrlZ                                                 : boolean;

key0, key1, key2, key3, key4, key5, key6, key7, key8, key9   : boolean;

Alt0, Alt1, Alt2, Alt3, Alt4, Alt5, Alt6, Alt7, Alt8, Alt9   : boolean;

F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12            : boolean;

AltF1, AltF2, AltF3, AltF4, AltF5, AltF6, AltF7, AltF8,
AltF9, AltF10, AltF11, AltF12                                : boolean;

ShiftF1, ShiftF2, ShiftF3, ShiftF4, ShiftF5, ShiftF6,
ShiftF7, ShiftF8,  ShiftF9, ShiftF10, ShiftF11, ShiftF12     : boolean;

CtrlF1, CtrlF2, CtrlF3, CtrlF4, CtrlF5, CtrlF6, CtrlF7,
CtrlF8, CtrlF9, CtrlF10, CtrlF11, CtrlF12                    : boolean;

Ins, Del, Home, En, PgUp, PgDn, CurUp, CurDn, CurLt, CurRt,

CtrlPgUp, CtrlPgDn                                           : boolean;

CtrlCurLt, CtrlCurRt                                         : boolean;

CtrlEn, CtrlHome                                             : boolean;

Minus, Plus, Mult, Di                                        : boolean;

Space, BkSp, esc, CR                                         : boolean;

funkey                                                       : boolean;

ch, cha                                                      : char;





implementation





procedure resetkeys;
 begin

 keyA:= false;
 keyB:= false;
 keyC:= false;
 keyD:= false;
 keyE:= false;
 keyF:= false;
 keyG:= false;
 keyH:= false;
 keyI:= false;
 keyJ:= false;
 keyK:= false;
 keyL:= false;
 keyM:= false;
 keyN:= false;
 keyO:= false;
 keyP:= false;
 keyQ:= false;
 keyR:= false;
 keyS:= false;
 keyT:= false;
 keyU:= false;
 keyV:= false;
 keyW:= false;
 keyX:= false;
 keyY:= false;
 keyZ:= false;

 AltA:= false;
 AltB:= false;
 AltC:= false;
 AltD:= false;
 AltE:= false;
 AltF:= false;
 AltG:= false;
 AltH:= false;
 AltI:= false;
 AltJ:= false;
 AltK:= false;
 AltL:= false;
 AltM:= false;
 AltN:= false;
 AltO:= false;
 AltP:= false;
 AltQ:= false;
 AltR:= false;
 AltS:= false;
 AltT:= false;
 AltU:= false;
 AltV:= false;
 AltW:= false;
 AltX:= false;
 AltY:= false;
 AltZ:= false;

 CtrlA:= false;
 CtrlB:= false;
 CtrlC:= false;
 CtrlD:= false;
 CtrlE:= false;
 CtrlF:= false;
 CtrlG:= false;
 CtrlH:= false;
 CtrlI:= false;
 CtrlJ:= false;
 CtrlK:= false;
 CtrlL:= false;
 CtrlM:= false;
 CtrlN:= false;
 CtrlO:= false;
 CtrlP:= false;
 CtrlQ:= false;
 CtrlR:= false;
 CtrlS:= false;
 CtrlT:= false;
 CtrlU:= false;
 CtrlV:= false;
 CtrlW:= false;
 CtrlX:= false;
 CtrlY:= false;
 CtrlZ:= false;

 key0:= false;
 key1:= false;
 key2:= false;
 key3:= false;
 key4:= false;
 key5:= false;
 key6:= false;
 key7:= false;
 key8:= false;
 key9:= false;

 Alt0:= false;
 Alt1:= false;
 Alt2:= false;
 Alt3:= false;
 Alt4:= false;
 Alt5:= false;
 Alt6:= false;
 Alt7:= false;
 Alt8:= false;
 Alt9:= false;

 F1:=  false;
 F2:=  false;
 F3:=  false;
 F4:=  false;
 F5:=  false;
 F6:=  false;
 F7:=  false;
 F8:=  false;
 F9:=  false;
 F10:= false;
 F11:= false;
 F12:= false;

 AltF1:=  false;
 AltF2:=  false;
 AltF3:=  false;
 AltF4:=  false;
 AltF5:=  false;
 AltF6:=  false;
 AltF7:=  false;
 AltF8:=  false;
 AltF9:=  false;
 AltF10:= false;
 AltF11:= false;
 AltF12:= false;

 ShiftF1:=  false;
 ShiftF2:=  false;
 ShiftF3:=  false;
 ShiftF4:=  false;
 ShiftF5:=  false;
 ShiftF6:=  false;
 ShiftF7:=  false;
 ShiftF8:=  false;
 ShiftF9:=  false;
 ShiftF10:= false;
 ShiftF11:= false;
 ShiftF12:= false;

 CtrlF1:=  false;
 CtrlF2:=  false;
 CtrlF3:=  false;
 CtrlF4:=  false;
 CtrlF5:=  false;
 CtrlF6:=  false;
 CtrlF7:=  false;
 CtrlF8:=  false;
 CtrlF9:=  false;
 CtrlF10:= false;
 CtrlF11:= false;
 CtrlF12:= false;

 ins:=  false;
 del:=  false;
 home:= false;
 en:=   false;

 Pgup:=  false;
 Pgdn:=  false;
 CurUp:= false;
 CurDn:= false;
 CurLt:= false;
 CurRt:= false;


 CtrlPgUp:=  false;
 CtrlPgDn:=  false;
 CtrlCurLt:= false;
 CtrlCurRt:= false;
 CtrlEn:=    false;
 CtrlHome:=  false;


 minus:= false;
 plus:=  false;
 mult:=  false;
 di:=    false;

 Space:= false;
 BkSp:=  false;
 esc:=   false;
 CR:=    false;

 funkey:= false;

 ch:= #0;

 end;   { of procedure resetkeys }





procedure presskey;
 begin
 resetkeys;
 cha:= readkey;

 ch:= upcase(cha);

 case ch of

           #65 : keyA:= true;
           #66 : keyB:= true;
           #67 : keyC:= true;
           #68 : keyD:= true;
           #69 : keyE:= true;
           #70 : keyF:= true;
           #71 : keyG:= true;
           #72 : keyH:= true;
           #73 : keyI:= true;
           #74 : keyJ:= true;
           #75 : keyK:= true;
           #76 : keyL:= true;
           #77 : keyM:= true;
           #78 : keyN:= true;
           #79 : keyO:= true;
           #80 : keyP:= true;
           #81 : keyQ:= true;
           #82 : keyR:= true;
           #83 : keyS:= true;
           #84 : keyT:= true;
           #85 : keyU:= true;
           #86 : keyV:= true;
           #87 : keyW:= true;
           #88 : keyX:= true;
           #89 : keyY:= true;
           #90 : keyZ:= true;

            #1 : CtrlA:= true;
            #2 : CtrlB:= true;
            #3 : CtrlC:= true;
            #4 : CtrlD:= true;
            #5 : CtrlE:= true;
            #6 : CtrlF:= true;
            #7 : CtrlG:= true;
            #8 : begin
                 CtrlH:= true;
                 BkSp:= true;
                 end;
            #9 : CtrlI:= true;
           #10 : CtrlJ:= true;
           #11 : CtrlK:= true;
           #12 : CtrlL:= true;
           #13 : begin
                 CtrlM:= true;
                 CR:= true;
                 end;
           #14 : CtrlN:= true;
           #15 : CtrlO:= true;
           #16 : CtrlP:= true;
           #17 : CtrlQ:= true;
           #18 : CtrlR:= true;
           #19 : CtrlS:= true;
           #20 : CtrlT:= true;
           #21 : CtrlU:= true;
           #22 : CtrlV:= true;
           #23 : CtrlW:= true;
           #24 : CtrlX:= true;
           #25 : CtrlY:= true;
           #26 : CtrlZ:= true;

           #48 : key0:= true;
           #49 : key1:= true;
           #50 : key2:= true;
           #51 : key3:= true;
           #52 : key4:= true;
           #53 : key5:= true;
           #54 : key6:= true;
           #55 : key7:= true;
           #56 : key8:= true;
           #57 : key9:= true;

           #45 : Minus:= true;
           #43 : Plus:= true;
           #42 : Mult:= true;
           #47 : Di:= true;

           #32 : Space:= true;
           #27 : esc:= true;

           end;


 if ch= #0 then

  begin
   funkey:= true;
   ch:= readkey;

   case ch of

              #16 : AltQ:= true;
              #17 : AltW:= true;
              #18 : AltE:= true;
              #19 : AltR:= true;
              #20 : AltT:= true;
              #21 : AltY:= true;
              #22 : AltU:= true;
              #23 : AltI:= true;
              #24 : AltO:= true;
              #25 : AltP:= true;
              #30 : AltA:= true;
              #31 : AltS:= true;
              #32 : AltD:= true;
              #33 : AltF:= true;
              #34 : AltG:= true;
              #35 : AltH:= true;
              #36 : AltJ:= true;
              #37 : AltK:= true;
              #38 : AltL:= true;
              #44 : AltZ:= true;
              #45 : AltX:= true;
              #46 : AltC:= true;
              #47 : AltV:= true;
              #48 : AltB:= true;
              #49 : AltN:= true;
              #50 : AltM:= true;

             #120 : Alt1:= true;
             #121 : Alt2:= true;
             #122 : Alt3:= true;
             #123 : Alt4:= true;
             #124 : Alt5:= true;
             #125 : Alt6:= true;
             #126 : Alt7:= true;
             #127 : Alt8:= true;
             #128 : Alt9:= true;
             #129 : Alt0:= true;

              #59 : F1:= true;
              #60 : F2:= true;
              #61 : F3:= true;
              #62 : F4:= true;
              #63 : F5:= true;
              #64 : F6:= true;
              #65 : F7:= true;
              #66 : F8:= true;
              #67 : F9:= true;
              #68 : F10:= true;
             #133 : F11:= true;
             #134 : F12:= true;

             #104 : AltF1:= true;
             #105 : AltF2:= true;
             #106 : AltF3:= true;
             #107 : AltF4:= true;
             #108 : AltF5:= true;
             #109 : AltF6:= true;
             #110 : AltF7:= true;
             #111 : AltF8:= true;
             #112 : AltF9:= true;
             #113 : AltF10:= true;
             #139 : AltF11:= true;
             #140 : AltF12:= true;

              #84 : ShiftF1:= true;
              #85 : ShiftF2:= true;
              #86 : ShiftF3:= true;
              #87 : ShiftF4:= true;
              #88 : ShiftF5:= true;
              #89 : ShiftF6:= true;
              #90 : ShiftF7:= true;
              #91 : ShiftF8:= true;
              #92 : ShiftF9:= true;
              #93 : ShiftF10:= true;
             #135 : ShiftF11:= true;
             #136 : ShiftF12:= true;

              #94 : CtrlF1:= true;
              #95 : CtrlF2:= true;
              #96 : CtrlF3:= true;
              #97 : CtrlF4:= true;
              #98 : CtrlF5:= true;
              #99 : CtrlF6:= true;
             #100 : CtrlF7:= true;
             #101 : CtrlF8:= true;
             #102 : CtrlF9:= true;
             #103 : CtrlF10:= true;
             #137 : CtrlF11:= true;
             #138 : CtrlF12:= true;

              #82 : Ins:= true;
              #83 : Del:= true;
              #71 : Home:= true;
              #79 : En:= true;

              #73 : PgUp:= true;
              #81 : PgDn:= true;
              #72 : CurUp:= true;
              #80 : CurDn:= true;
              #75 : CurLt:= true;
              #77 : CurRt:= true;

             #132 : CtrlPgUp:= true;
             #118 : CtrlPgDn:= true;
             #115 : CtrlCurLt:= true;
             #116 : CtrlCurRt:= true;
             #117 : CtrlEn:= true;
             #119 : CtrlHome:= true;

             end;

  end;

 end;   { of procedure presskey }










begin

resetkeys;

end.   { of unit KEY }
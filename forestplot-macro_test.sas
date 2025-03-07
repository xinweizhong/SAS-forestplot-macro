
**************************示例********************************************;
Options notes nomprint nosymbolgen nomlogic nofmterr nosource nosource2 missing=' ' noquotelenmax linesize=max noBYLINE;
dm "output;clear;log;clear;odsresult;clear;";
proc delete data=_all_; run;
%macro rootpath;
%global program_path program_name;
%if %symexist(_SASPROGRAMFILE) %then %let _fpath=%qsysfunc(compress(&_SASPROGRAMFILE,"'"));
	%else %let _fpath=%sysget(SAS_EXECFILEPATH);
%let program_path=%sysfunc(prxchange(s/(.*)\\.*/\1/,-1,%upcase(&_fpath.)));
%let program_name=%scan(&_fpath., -2, .\);
%put NOTE: ----[program_path = &program_path.]----;
%put NOTE: ----[program_name = &program_name.]----;
%mend rootpath;
%rootpath;

%inc "&program_path.\forestplot-macro.sas";
/*Output styles settings*/
options nodate nonumber nobyline;
ods path work.testtemp(update) sasuser.templat(update) sashelp.tmplmst(read);

Proc template;
  define style trial;
    parent=styles.rtf;
    style table from output /
    background=_undef_
    rules=groups
    frame=void
    cellpadding=1pt;
  style header from header /
    background=_undef_
    protectspecialchars=off;
  style rowheader from rowheader /
    background=_undef_;

  replace fonts /
    'titlefont5' = ("courier new",9pt)
    'titlefont4' = ("courier new",9pt)
    'titlefont3' = ("courier new",9pt)
    'titlefont2' = ("courier new",9pt)
    'titlefont'  = ("courier new",9pt)
    'strongfont' = ("courier new",9pt)
    'emphasisfont' = ("courier new",9pt)
    'fixedemphasisfont' = ("courier new",9pt)
    'fixedstrongfont' = ("courier new",9pt)
    'fixedheadingfont' = ("courier new",9pt)
    'batchfixedfont' = ("courier new",9pt)
    'fixedfont' = ("courier new",9pt)
    'headingemphasisfont' = ("courier new",9pt)
    'headingfont' = ("courier new",9pt)
    'docfont' = ("courier new",9pt);

  style body from body /
    leftmargin=0.1in
    rightmargin=0.1in
    topmargin=0.1in
    bottommargin=0.1in;

  class graphwalls / 
            frameborder=off;
   end;
run;

options source source2;
**************************外部导入annotate dataset*****************************;
/*************************
注意： 
图中第1列对应id值为myid1；
图中第2列对应id值为myid2；
图中第3列对应id值为myid3；
..........
图中第n列对应id值为myidn；

layout lattice最外部的id值为myid ;
**************************/

********************************************************************;
******示例1数据******;
proc import datafile="&program_path.\forestplot.xlsx"
  out=forestplot
  dbms=excel replace;
  sheet ='sheet1';
  getnames=yes;
quit;
******示例2数据******;
proc import datafile="&program_path.\forestplot.xlsx"
  out=forestplot1
  dbms=excel replace;
  sheet ='sheet3';
  getnames=yes;
quit;
******示例3数据******;
proc import datafile="&program_path.\forestplot.xlsx"
  out=forestplot2
  dbms=excel replace;
  sheet ='sheet2';
  getnames=yes;
quit;
******示例4数据******;
proc import datafile="&program_path.\forestplot.xlsx"
  out=forestplot3
  dbms=excel replace;
  sheet ='sheet4';
  getnames=yes;
quit;

******示例5数据******;
proc import datafile="&program_path.\forestplot.xlsx"
  out=forestplot4
  dbms=excel replace;
  sheet ='sheet5';
  getnames=yes;
quit;


**************外部导入annotate dataset*********************;
%let fcol=4;
data anno_arrow;
	length id function x1space y1space x2space y2space linecolor shape direction $20.;
	retain function 'arrow' x1space "datavalue" y1space "layoutpercent" x2space "datavalue" y2space "layoutpercent" 
		linecolor "black"; 
	linethickness=0.4; shape='barbed'; id="myid&fcol.";
	x1=0.2; y1=-1.5; x2=0.9; y2=-1.5; direction='in'; output;
	x1=1.1; y1=-1.5; x2=1.9; y2=-1.5; direction='out'; output;
run;
data anno_text;
	length id function x1space y1space anchor textcolor $20. label $200.;
	retain function 'text' x1space "datavalue" y1space "layoutpercent" anchor 'left' textcolor "black";
	textstyle='normal'; textweight='normal'; width=200;
	textsize=8;  textfont='Arial'; id="myid&fcol.";
	x1=0.2; y1=-5; label='Favours Group 01'; output;
	x1=1.2; y1=-5; label='Favours Group 02'; output;
run;
data anno_;
	set anno_arrow anno_text;
run;

data anno_line;
	length id function x1space y1space x2space y2space linecolor  $20.;
	retain function 'line' x1space "layoutpercent" y1space "datavalue" x2space "layoutpercent" y2space "datavalue" 
		linecolor "black"; 
	linethickness=2; 
	id="myid2"; x1=1; y1=1.5; x2=100; y2=y1; output;
	id="myid3"; output;
	id="myid4"; output;
	id="myid5"; output;
run;

ods _all_ close;
title;footnote;
ods results on;
goption device=pdf;
options topmargin=0.1in bottommargin=0.1in leftmargin=0.1in rightmargin=0.1in;
options orientation=landscape nodate nonumber;
ods pdf file="&program_path.\forestplot_template.pdf"  style=trial nogtitle nogfoot ;

ods graphics on; 
ods graphics /reset  noborder maxlegendarea=55  outputfmt =pdf height = 7.2 in width = 10.6in  attrpriority=none;
ods escapechar='^';

data forestplot;
	set forestplot;
	if strip(subtitle)='Category#' then subtitle="Category[ ^{Unicode '03C7'x}^{Unicode '2122'x}^{Unicode '00AE'x} ]";
	LIMITLOWER=0.4; LIMITUPPER=1.5;
run;
************************* 示例1 *******************************************;
%let columnweights=%str(0.20 0.14 0.14 0.38 0.14);
%let xaxisls=%str(0 0.5 1 2);
%ForestPlot(indat=forestplot
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=50/JUST=C/ERRORBARCAPSCALE=0.1/markercolor=red/ERRORBARCOLOR=blue|
T=or_ci95/X=50/JUST=C|)
					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=
					,refline=%str(COLOR=red)
					,logaxis= 
					,fontsize=
					,textfont=%str(Arial/simsun)
					,width=
					,colorbands=data
/*					,bandplot=%str(LIMITLOWER=0.4/LIMITUPPER=1.5/fillcolor=red)*/
					,bandplot=%str(LIMITLOWER=LIMITLOWER/LIMITUPPER=LIMITUPPER/fillcolor=red)
					,f_xoffset=%str(0|0)
					,add_annods=%str(anno_)
					,debug=zxw);

****************************** 示例5 **************************************;
%let columnweights=%str(0.14 0.10 0.10 0.23 0.10 0.23 0.10 );
%let xaxisls=%str(0 0.5 1 2|0 0.5 1 2);
%ForestPlot(indat=forestplot4
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
F=lowerci1#oddsratio1#upperci1/title=Odds Ratio (95% CI)/X=50/JUST=C/ERRORBARCAPSHAPE=none/markercolor=red/ERRORBARCOLOR=blue|
T=or_ci951/X=50/JUST=C|
F=lowerci2#oddsratio2#upperci2/title=Odds Ratio (95% CI)/X=50/JUST=C/ERRORBARCAPSHAPE=none/markercolor=red/ERRORBARCOLOR=blue|
T=or_ci952/X=50/JUST=C|)
					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=
					,refline=%str(COLOR=red|COLOR=blue)
					,logaxis= 
					,fontsize=
					,textfont=%str(Arial/simsun)
					,width=
					,colorbands=even
					,bandplot=%str(LIMITLOWER=0.4/LIMITUPPER=1.5/fillcolor=red|LIMITLOWER=0.2/LIMITUPPER=1/fillcolor=green)
/*					,bandplot=%str(LIMITLOWER=LIMITLOWER/LIMITUPPER=LIMITUPPER/fillcolor=red)*/
					,f_xoffset=%str(0|0)
					,add_annods=%str()
					,debug=zxw);

****************************** 示例2 **************************************;
%let columnweights=%str(0.12 0.12 0.12 0.12 0.12 0.28 0.12);
%let xaxisls=%str(0.1 1 10);
%ForestPlot(indat=forestplot1
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=median_ci1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
T=median_ci2/X=33/JUST=C|
F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=50/JUST=C|
T=or_ci95/X=50/JUST=C|)

					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=
					,refline=
					,logaxis= 10/*BASE=10 | 2 | E */
					,fontsize=
					,textfont=%str(Arial)
					,width=
					,colorbands=ODD
					,f_xoffset=%str(0|0)
					,add_annods=%str(anno_line)
					,debug=1);


************************* 示例4 *******************************************;
%let columnweights=%str(0.12 0.12 0.12 0.12 0.12 0.28 0.12);
%let xaxisls=%str(0.1 1 10);
%ForestPlot(indat=forestplot3
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=median_ci1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
T=median_ci2/X=33/JUST=C|
F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=50/JUST=C|
T=or_ci95/X=50/JUST=C|)

,mergeheader=%str(
    GROUP1/#/[c]GROUP 01/[c]TOTAL GROUP|
median_ci1/#/[c]GROUP 01/[c]TOTAL GROUP|
    GROUP2/#/[c]GROUP 02/[c]TOTAL GROUP|
median_ci2/#/[c]GROUP 02/[c]TOTAL GROUP|)
					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=
					,refline=
					,logaxis= 10/*BASE=10 | 2 | E */
					,fontsize=
					,textfont=%str(Arial)
					,width=
					,header_colorbands=%str(111)
					,colorbands=data
					,f_xoffset=%str(0|0)
					,add_annods=%str()
					,debug=0);

ods pdf close;
ods listing;

******************************************************;
title;footnote;
goption device=png;
options topmargin=0.1in bottommargin=0.1in leftmargin=0.1in rightmargin=0.1in;
options orientation=landscape nodate nonumber;
ods rtf file="&program_path.\forestplot_template.rtf"  style=trial nogtitle nogfoot ;

ods graphics on; 
ods graphics /reset  noborder maxlegendarea=55  outputfmt =png height = 7.2 in width = 10.6in  attrpriority=none;
ods escapechar='^';

************************* 示例3 *******************************************;
%let columnweights=%str(0.20 0.14 0.14 0.38 0.14);
%let xaxisls=%str(0 0.5 1 2);
%ForestPlot(indat=forestplot2
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=50/JUST=C|
T=or_ci95/X=50/JUST=C|)
					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=%str(bottom=0.7cm)
					,refline=
					,logaxis= 
					,fontsize=
					,textfont=%str(SIMSUN/Arial)
					,width=
					,colorbands=even
					,add_annods=%str(anno_)
					,debug=0);

********************************;
************************* 示例4 *******************************************;
%let columnweights=%str(0.12 0.12 0.12 0.12 0.12 0.28 0.12);
%let xaxisls=%str(0.1 1 10);
%ForestPlot(indat=forestplot3
					,col_info=%str(
T=subtitle/X=0|
T=GROUP1/X=33/JUST=C|
T=median_ci1/X=33/JUST=C|
T=GROUP2/X=33/JUST=C|
T=median_ci2/X=33/JUST=C|
F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=50/JUST=C|
T=or_ci95/X=50/JUST=C|)

,mergeheader=%str(
GROUP1/#/[l]GROUP 01/[c]TOTAL GROUP|
median_ci1/#/[l]GROUP 01/[c]TOTAL GROUP|
GROUP2/#/[c]GROUP 02/[c]TOTAL GROUP|
median_ci2/#/[c]GROUP 02/[c]TOTAL GROUP|)
					,columnweights=%str(&columnweights.)
					,xaxisls=%str(&xaxisls.)
					,xaxislsc=
					,pad=
					,refline=
					,logaxis= 10/*BASE=10 | 2 | E */
					,fontsize=
					,textfont=%str(Arial)
					,width=
					,colorbands=even
					,f_xoffset=%str(0|0)
					,add_annods=%str()
					,debug=zxw);


ods rtf close;
ods listing;

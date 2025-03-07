
***********************************************************************************************;
*%backup_program;

%macro ForestPlot(indat=
					,col_info=
					,mergeheader=
					,columnweights=
					,xaxisls= 
					,xaxislsc= 
					,pad=
					,refline= 
					,logaxis= 
					,fontsize=
					,textfont=
					,width=
					,add_annods=
					,f_xoffset= 
					,columngutter=
					,colorbands=data
					,header_colorbands=
					,bandoffset= 
					,bandplot= 
					,colorbandsattrs=
					,debug=);

*****parameter control****;
%if %length(&indat.)=0 %then %do;
	%put ERR%str()OR: Parameter[indat] uninitialized, please check!!;
	%return;
%end;
%if %sysfunc(exist(&indat.))=0 %then %do;
	%put ERR%str()OR: DataSet[indat = &indat.] no exist, please check!!;
	%return;
%end;


%if %length(&col_info.)>0 %then %do;
data __tmp;
length cat attrib $20.;
cat='T'; ord=1; attrib='VAR1'; output;
cat='T'; ord=2; attrib='X'; output;
cat='T'; ord=3; attrib='JUST'; output;
cat='F'; ord=1; attrib='VAR1'; output;
cat='F'; ord=2; attrib='VAR2'; output;
cat='F'; ord=3; attrib='VAR3'; output;
cat='F'; ord=4; attrib='TITLE'; output;
cat='F'; ord=5; attrib='X'; output;
cat='F'; ord=6; attrib='JUST'; output;
cat='F'; ord=7; attrib='ERRORBARCAPSHAPE'; output;
cat='F'; ord=8; attrib='MARKERCOLOR'; output;
cat='F'; ord=9; attrib='MARKERSYMBOL'; output;
cat='F'; ord=10; attrib='MARKERSIZE'; output;
cat='F'; ord=11; attrib='MARKERWEIGHT'; output;
cat='F'; ord=12; attrib='ERRORBARCAPSCALE'; output;
cat='F'; ord=13; attrib='ERRORBARCOLOR'; output;
cat='F'; ord=14; attrib='ERRORBARPATTERN'; output;
cat='F'; ord=15; attrib='ERRORBARTHICKNESS'; output;
run;

/*%let col_info=%str(T=subtitle/X=1|T=GROUP11/X=1|T=GROUP2/X=1|F=lowerci#oddsratio#upperci/title=Odds Ratio (95% CI)/X=1|T=or_ci95/X=1|);*/
%let coltperr=0;
data __col_info(drop=attribval1);
	length col_info $1000. subcol $500. cat attrib $20. attribval attribval1 $200.;
	col_info=tranwrd(tranwrd(tranwrd("&col_info.",'%/','$#@'),'%\','*#@'),'||','|');
	if substr(col_info,lengthn(col_info),1)='|' then col_info=substr(col_info,1,lengthn(col_info)-1);
	do coln=1 to (count(col_info,'|')+1);
		subcol=scan(col_info,coln,'|'); 
		if substr(subcol,lengthn(subcol),1)='/' then subcol=substr(subcol,1,lengthn(subcol)-1);
		if substr(upcase(compress(subcol)),1,2) in ('T=','F=') then do;
			cat=substr(upcase(compress(subcol)),1,1);
			do ord=1 to (count(subcol,'/')+1);
				attribval=strip(scan(subcol,ord,'/'));
				attrib=strip(upcase(scan(attribval,1,'=')));
				attribval=strip(scan(attribval,2,'='));
				if ord=1 and attrib in ('T','F') then do;
					if strip(attrib)='T' then attrib='VAR1';
					if strip(attrib)='F' then do;
						attribval1=attribval;
						attrib='VAR1'; attribval=strip(scan(attribval1,1,'#'));output;
						attrib='VAR2'; attribval=strip(scan(attribval1,2,'#'));output;
						attrib='VAR3'; attribval=strip(scan(attribval1,3,'#'));
					end;
				end;
				output;
			end;
		end;else do;
			call symputx('coltperr','1');
			put "ERR" "OR:" subcol= "column type not in [T=,F=], please check!";
		end;
	end;
run;
%if &coltperr.>0 %then %do;
	%put ERR%str()OR: column type not in [T=,F=], please check!!;
	%return;
%end;
proc sql undo_policy=none;
	create table __tmp1 as select a.coln,a.cat,b.ord,b.attrib
		from (select distinct coln,cat from __col_info) as a,__tmp as b where a.cat=b.cat
		order by a.coln,b.ord;
	create table __col_info as select a.coln,b.coln as coln1,a.cat,b.cat as cat1,a.ord,b.ord as ord1
		,a.attrib,b.attrib as attrib1,b.attribval
		from __tmp1 as a
		left join __col_info as b on a.coln=b.coln and a.cat=b.cat and a.attrib=b.attrib
		order by a.coln,a.ord;
quit; 
%let varerr=0;
%let _fnum=0; %let _fcol1=0;
data __col_info;
	set __col_info(where=(^missing(coln))) end=last;
	by coln ord;
	retain _fnum 0;
	attribval=tranwrd(tranwrd(attribval,'$#@','/'),'*#@','\');
	array _ms(*)  cat attrib;
	array _ms1(*)  cat1 attrib1;
	do i=1 to dim(_ms);
		if missing(_ms(i)) then _ms(i)=_ms1(i);
	end;
	if coln=. then coln=coln1;
	if strip(attrib)='VAR1' then call symputx(cats('col',coln,'var1'),attribval);
	if strip(cat)='F' then do;
		if strip(attrib)='VAR2' then call symputx(cats('col',coln,'var2'),attribval);
		if strip(attrib)='VAR3' then call symputx(cats('col',coln,'var3'),attribval);
		if strip(attrib)='TITLE' then call symputx(cats('col',coln,'title'),attribval);
		if ord=1 then _fnum+1;
		call symputx('_fnum',put(_fnum,best.));
		call symputx(cats('_fcol',_fnum),put(coln,best.));
	end;
	if strip(attrib)='X' then do;
		if strip(attribval)='' then attribval='0';
		call symputx(cats('col',coln,'x'),attribval);
	end;
	if strip(attrib)='JUST' then do;
		if strip(attribval)='' then attribval='LEFT';
		call symputx(cats('col',coln,'just'),attribval);
	end;
	if strip(attrib)='ERRORBARCAPSHAPE' then do;
		if strip(attribval)='' then attribval='SERIF';
		call symputx(cats('ERRORBARCAPSHAPE'),attribval);
	end;
	if strip(attrib)='MARKERCOLOR' then do;
		if strip(attribval)='' then attribval='CX000000';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='MARKERSYMBOL' then do;
		if strip(attribval)='' then attribval='CIRCLEFILLED';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='MARKERSIZE' then do;
		if strip(attribval)='' then attribval='10';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='MARKERWEIGHT' then do;
		if strip(attribval)='' then attribval='';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='ERRORBARCAPSCALE' then do;
		if strip(attribval)='' then attribval='1';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='ERRORBARCOLOR' then do;
		if strip(attribval)='' then attribval='CX000000';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='ERRORBARPATTERN' then do;
		if strip(attribval)='' then attribval='1';
		call symputx(strip(attrib),attribval);
	end;
	if strip(attrib)='ERRORBARTHICKNESS' then do;
		if strip(attribval)='' then attribval='2';
		call symputx(strip(attrib),attribval);
	end;
	if last.coln then call symputx(cats('col',coln,'type'),cat);
	if last then call symputx('columnsn',put(coln,best.));
	****check variable****;
	if index(attrib,'VAR') then do;
		dsid=open("&indat.", "i");
		num=varnum(dsid,attribval);
		if num<1 then do;
			call symputx('varerr','1');
			put "ERR" "OR:" coln= attrib= attribval= "varibale not exist, please check!";
		end;
		rc=close(dsid);
	end;
run;
%if &varerr.>0 %then %do;
	%put ERR%str()OR: Some varibale not exist in dataset[&indat.], please check!!;
	%return;
%end;
/*%put &col1var1. | &col1x. | &col4var1. | &col4var2. | &col4var3. | &col4x. | &col4title.;*/
%end; %else %do;
	%put ERR%str()OR: Parameter[col_info] uninitialized, please check!!;
	%return;
%end;

%if %length(&pad.)=0 %then %let pad=%str(bottom=0.1cm);
%if %length(&refline.)=0 %then %let refline=%str(X=1/DATATRANSPARENCY=0/color=black/PATTERN=ShortDash/THICKNESS=1);

data __tmp_fcol;
	%do _i=1 %to &_fnum.;
		coln=&&_fcol&_i.; output;
	%end;
run;
data __tmp_refl;
length cat attrib $20.;
cat='R'; ord=1; attrib='X'; output;
cat='R'; ord=2; attrib='DATATRANSPARENCY'; output;
cat='R'; ord=3; attrib='COLOR'; output;
cat='R'; ord=4; attrib='PATTERN'; output;
cat='R'; ord=5; attrib='THICKNESS'; output;
cat='R'; ord=6; attrib='OFFSET'; output;
run;

data __refline;
	length refline_info $3000. subrefl $1000. subcol $500. attrib value $50.;
	refline_info="&refline.";
	do n=1 to (count(refline_info,'|')+1);
		coln=input(symget(cats('_fcol',n)),best.);
		subrefl=strip(scan(refline_info,n,'|'));
		do ord=1 to (count(subrefl,'/')+1);
			subcol=strip(scan(subrefl,ord,'/'));
			attrib=strip(upcase(scan(subcol,1,'=')));
			value=strip(scan(subcol,2,'='));	
			output;
		end;
	end;
run;
proc sql undo_policy=none;
	create table __tmp_refl1 as select a.coln,b.cat,b.ord,b.attrib
		from (select distinct coln from __tmp_fcol) as a,__tmp_refl as b
		order by a.coln,b.ord;

	create table __refline as select a.coln,a.cat,a.ord,a.attrib,b.attrib as attrib1,b.value
		from __tmp_refl1 as a left join __refline as b
		on a.coln=b.coln and a.attrib=b.attrib
		order by a.coln,a.ord;
quit;

data __refline;
	set __refline;
	if strip(attrib)='X' then do;
		if value='' then value='1';
		call symputx(cats('refline',coln,'_x'),value);
	end;
	if strip(attrib)='DATATRANSPARENCY' then do;
		if value='' then value='0';
		call symputx(cats('refline',coln,'_TRANSPARENCY'),value);
	end; 
	if strip(attrib)='COLOR' then do;
		if value='' then value='black';
		call symputx(cats('refline',coln,'_COLOR'),value);
	end;
	if strip(attrib)='PATTERN' then do;
		if value='' then value='ShortDash';
		call symputx(cats('refline',coln,'_PATTERN'),value);
	end;
	if strip(attrib)='THICKNESS' then do;
		if value='' then value='1';
		call symputx(cats('refline',coln,'_THICKNESS'),value);
	end;
	if strip(attrib)='OFFSET' then do;
		if value='' then value='0.5';
		call symputx(cats('refline',coln,'_OFFSET'),value);
	end;
run;

%if %length(&fontsize.)=0 %then %let fontsize=9;
%if %length(&textfont.)=0 %then %let textfont=%str(SimSun/Arial);
%if %length(&width.)=0 %then %let width=500;
%if %length(&add_annods.)>0 %then %do;
	%if %sysfunc(exist(&add_annods.))=0 %then %do;
		%put ERR%str()OR: DataSet[indat = &add_annods.] no exist, please check!!;
		%return;
	%end;
%end;

%let _xoffsetmin=0.01; %let _xoffsetmax=0.01;
%if %length(&f_xoffset.)>0 %then %do;
	%let _xoffsetmin=%scan(&f_xoffset.,1,#|); 
	%let _xoffsetmax=%scan(&f_xoffset.,2,#|); 
%end;
%if %length(&_xoffsetmin.)=0 %then %let _xoffsetmin=0.01; 
%if %length(&_xoffsetmax.)=0 %then %let _xoffsetmax=0.01;
%if %length(&columngutter.)=0 %then %let columngutter=0pt;
%let _colorbands=;
%if %length(&colorbands.)>0 %then %do;
	%let _colorbands=%upcase(&colorbands.);
	%if "&_colorbands."^="EVEN" and "&_colorbands."^="ODD" and "&_colorbands."^="DATA" %then %let _colorbands=;
	%if "&_colorbands."="EVEN" %then %let _bandsx=1;
		%else %if "&_colorbands."="ODD" %then %let _bandsx=0;
		%else %if "&_colorbands."="DATA" %then %let _bandsx=1;
%end;
%if %length(&bandoffset.)=0 %then %let bandoffset=0;
/*%put &=_xoffsetmin &=_xoffsetmax &=_colorbands.;*/
%if %length(&colorbandsattrs.)=0 %then %let colorbandsattrs=%str(FILLTRANSPARENCY=0.6; FILLCOLOR="lightgray";) ;

************ mergeheader *********************;
%let _addheadern=0;
%if %length(&mergeheader.)>0 %then %do;
data __colw;
	length cols $200. colw $20.;
	cols=strip(compbl("&columnweights."));
	do coln=1 to count(strip(cols),' ')+1;
		colw=scan(cols,coln,' '); 
		call symputx(cats('col',coln,'wei'),colw); output;
	end;
run;
%macro deal_xaxis(colls=,just=,macro_var=);
%let colls=%sysfunc(tranwrd(&colls.,#,%str(,)));
data __colwei;
	min_col=min(&colls.);
	max_col=max(&colls.);
	do i=1 to max_col;
		colw=input(symget(cats('col',i,'wei')),best.); output;
	end;
run;
%global min_col max_col &macro_var.;
%let min_col=0; %let max_col=;  %let &macro_var.=; 
data __colwei;
	set __colwei;
	retain colws 0;
	colws=sum(colws,colw);
	if i=(min_col-1) then call symputx('min_col',put(colws,best.));
	if i=max_col then call symputx('max_col',put(colws,best.));
run;
%let just=%upcase(&just.);
%if "&just."="L" or "&just."="LEFT" %then %let &macro_var.=&min_col.;
	%else %if "&just."="R" or "&just."="RIGHT" %then %let &macro_var.=&max_col.;
	%else %if "&just."="C" or "&just."="CENTER" %then %let &macro_var.=%sysevalf(&min_col.+%sysevalf(%sysevalf(&max_col.-&min_col.)/2));
/*%put &=just. @ &=min_col. | &=max_col. | &macro_var.= &&&macro_var.;*/
%mend deal_xaxis;

data __mergeheader;
	length header_info $1000. subcol $500. attrib $20. value $200.;
	header_info=tranwrd(tranwrd(tranwrd("&mergeheader.",'%/','$#@'),'%\','*#@'),'||','|');
	if substr(header_info,lengthn(header_info),1)='|' then header_info=substr(header_info,1,lengthn(header_info)-1);
	do coln=1 to (count(header_info,'|')+1);
		subcol=strip(scan(header_info,coln,'|')); 
		if substr(subcol,lengthn(subcol),1)='/' then subcol=substr(subcol,1,lengthn(subcol)-1);
		do ord=1 to (count(subcol,'/')+1);
			value=strip(scan(subcol,ord,'/'));
			value=tranwrd(tranwrd(value,'$#@','/'),'*#@','\');
			if ord=1 then attrib='VAR';
				else if ord=2 then do; attrib='LABEL'; if strip(value) in ('#','') then value=''; end;
				else if ord>2 then attrib=cats('MERGEHEADER',ord-2);	
			output;
		end;
	end;
	drop header_info;
run;
proc transpose data=__mergeheader out=__mergeheader1 prefix=mergeheader;
	var value;
	by coln;
	id ord;
run;
%let _header1nonull=0;
proc sql undo_policy=none noprint;
	select max(ord) into: _mgheadern trimmed from __mergeheader;
	create table __mergeheader1 as select a.*,b.coln,c.colw
		from __mergeheader1(drop=coln) as a
		left join __col_info(where=(prxmatch('#(VAR)(\d+)#',attrib))) as b 
			on strip(upcase(a.mergeheader1))=strip(upcase(b.attribval))
		left join __colw as c on b.coln=c.coln
		order by %do _h=&_mgheadern. %to 3 %by -1; a.mergeheader&_h.,%end; b.coln;
	select count(*) into: _header1nonull trimmed from __mergeheader1 where strip(mergeheader2) not in ('#','');
quit; 
%let _addheadern=%eval(&_mgheadern.-1);
%if &_header1nonull.<1 %then %let _addheadern=%eval(&_addheadern.-1);
/*%put &=_header1nonull. &=_addheadern.;*/
data __mergeheader1;
	set __mergeheader1 end=last;
	call symputx(cats('mgheader_coln',_n_),put(coln,best.));
	if last then call symputx('_mgtotcln',put(_n_,best.));
	%do _h=&_mgheadern. %to 3 %by -1;
		length anchor&_h. $50.;
		anchor&_h.='LEFT';
		id1=prxparse('#(\[)(\w+)(\])(\S+|\s+)#');
		if prxmatch(id1,mergeheader&_h.) then do;
			anchor&_h.=strip(upcase(prxposn(id1, 2, mergeheader&_h.)));
			if strip(anchor&_h.) in ('L','LEFT') then anchor&_h.='LEFT';
			if strip(anchor&_h.)='C' then anchor&_h.='CENTER';
			if strip(anchor&_h.)='R' then anchor&_h.='RIGHT';
			mergeheader&_h.=strip(scan(mergeheader&_h.,2,']'));
		end;
	%end; 
run;
proc sort data=__mergeheader1;
	by %do _h=&_mgheadern. %to 3 %by -1; mergeheader&_h. %end; coln;
run;
data %do _h=&_mgheadern. %to 3 %by -1; __mergeheader&_h. %end;;
 	set __mergeheader1;
	by %do _h=&_mgheadern. %to 3 %by -1; mergeheader&_h. %end; coln;
	
	%do _h=&_mgheadern. %to 3 %by -1;
		length colls&_h. $20.; retain colls&_h.;
		if first.mergeheader&_h. then colls&_h.='';
		colls&_h.=catx('#',colls&_h.,coln);
		if last.mergeheader&_h. and strip(mergeheader&_h.)>'' then output __mergeheader&_h.;
	%end;
run;
%do _h=&_mgheadern. %to 3 %by -1;
	data __mergeheader&_h.;
		set __mergeheader&_h. end=last;
		cln=_n_;
		call symputx(cats('colls',_n_),colls&_h.);
		call symputx(cats("hr&_h._just",_n_),anchor&_h.);
		call symputx(cats("hr&_h._header",_n_),mergeheader&_h.);
		if last then call symputx('_coln',cats(_n_));
	run;
	%do _l=1 %to &_coln.;
		%let _mvar=%str(hr&_h._axis&_l.); 
		%deal_xaxis(colls=%str(&&colls&_l.),just=%str(&&hr&_h._just&_l.),macro_var=%str(&_mvar.));
	%end;
%end;
%end;
	
%if %length(&debug.)=0 %then %let debug=0;
%if "&debug."^="zxw" %then %let debug=0;

%if "&debug."="0" %then %do;
	proc datasets nolist;
		delete __col_info __colw: __refline __bandplot;
	quit;
%end;

************************************************************************;
%let logstr=;
%do _i=1 %to &_fnum.;
	%let _coln=&&_fcol&_i.; 
	%let _nolog&_coln.=;
	%if &_i.=1 %then %let logstr=%str(^missing(&&col&_coln.var2.)); 
		%else %let logstr=%str(&logstr. or ^missing(&&col&_coln.var2.)); 
%end;

%let vartyperr=0;
%let dropliney=0;

%if &_addheadern.>0 and "&_colorbands."="DATA" %then %do;
%do _i=0 %to &_addheadern.;
	%let _colorbands&_i.=%substr(&header_colorbands.,%eval(&_i.+1),1); 
/*	%put  _colorbands&_i.=&&_colorbands&_i.;*/
%end;
%end;

data __final;
	set &indat. end=last;
	if _n_<1 then do; subind=0; boldind=0;  Colorbands=0; end;
	if subind=. then subind=0; 
	if boldind=. then boldind=0;
	if Colorbands=. then Colorbands=0;

	retain dropy 0;
	if &logstr. then dropy+1;
	__ord=_n_-1+&_addheadern.;
	
	array _col(&columnsn.) $2. __col1-__col&columnsn.;
	do i=1 to dim(_col);
		_col(i)='';
	end; drop i;

	%do _i=1 %to &_fnum.;
		%let _coln=&&_fcol&_i.;
		__refline&_coln.=&&refline&_coln._X.;
		if cmiss(&&col&_coln.var1,&&col&_coln.var2,&&col&_coln.var3)>0 then call missing(&&col&_coln.var1,&&col&_coln.var2,&&col&_coln.var3);

		if .<&&col&_coln.var1<=0 or .<&&col&_coln.var2<=0 or .<&&col&_coln.var3<=0 then call symputx("_nolog&_coln.",'1');
		%do _j=1 %to 3;
			if _n_=1 and strip(upcase(vtype(&&col&_coln.var&_j.)))^="N" then do;
				put "ERR" "OR: [&&col&_coln.var&_j.] no numeric variable!";
				call symputx('vartyperr','1');
			end;
		%end;
	%end;
	if missing(subind) then subind=0; 
	if missing(boldind) then boldind=0;
	if _n_=1 then call symputx('ymin',cats(__ord),'g');
	if last then call symputx('ymax',cats(__ord),'g');
run;
/*%put &=ymax.;*/
%if &vartyperr.>0 %then %do;
	%let varstr=;
	%do _i=1 %to &_fnum.; 
		%let _coln=&&_fcol&_i.;
		%let varstr=%str(&varstr. &&col&_coln.var1./&&col&_coln.var2./&&col&_coln.var3.);
	 %end;
	%put ERR%str()OR: Variable[  ] no numeric variable, please check!!;
	%return;
%end;
%do _i=1 %to &_fnum.;
	%let _coln=&&_fcol&_i.;
	%if &&_nolog&_coln.>0 %then %do;
		%put ERR%str()OR: Column Variable[&&col&_coln.var1./&&col&_coln.var2./&&col&_coln.var3.] ge 0, log conversion cannot be performed, please check!!;
		%return;
	%end;
%end;
data _null_;
	set __final(where=(dropy=1)) end=last;
	if last then call symputx('dropliney',cats(__ord));
run;
%put &=dropliney.;
%let __col1w=%scan(&columnweights,1,%str( ));
%if &__col1w.>1 %then %let __col1w=%sysevalf(&__col1w./100);

%if %length(&mergeheader.)>0 %then %do;
data __anno_col1;
	length id function x1space y1space anchor textcolor textweight textstyle textfont $50. label $300.;
	retain function 'text' x1space "layoutpercent" y1space "layoutpercent" anchor 'left' textcolor "black";
	set __mergeheader1(in=a where=(mergeheader2>'')) %do _h=&_mgheadern. %to 3 %by -1; __mergeheader&_h.(in=b&_h.) %end; ;
	textstyle='normal'; textweight='bold'; width=&width.;
	textsize=&fontsize.;  textfont="&textfont.";
 
	%if &_header1nonull.>0 %then %do;
		if a then do;
			y1=(&ymax.-&_addheadern.-0.5)/&ymax.*100;
			if strip(mergeheader2) not in ('#','') then do;
				id=cats("myid",coln); 
				label=strip(mergeheader2);
				x1=input(symget(cats("col",coln,'x')),best.);
				anchor=strip(symget(cats("col",coln,'just')));  
			end;
		end;
	%end; 
	%let cutn=%eval(&_mgheadern.-2);;
	%do _h=&_mgheadern. %to 3 %by -1;
		if b&_h. then do;
			id="myid1"; y1=(&ymax.-(&_addheadern.-&cutn.)-0.5)/&ymax.*100;
			anchor=strip(anchor&_h.);
			label=strip(mergeheader&_h.);
			x1=input(symget(cats("hr&_h._axis",cln)),best.)*100;
			x1=x1/(&__col1w.);
		end;
		%let cutn=%eval(&cutn.-1);
	%end;
	if label>'';
run;
data anno_mergeheader;
	length id function x1space y1space x2space y2space linecolor $50.;
	retain function 'line' x1space "layoutpercent" y1space "layoutpercent" x2space "layoutpercent" y2space "layoutpercent" 
		linecolor "black" ; 
	linethickness=2;
	set %do _h=&_mgheadern. %to 3 %by -1; __mergeheader&_h.(in=b&_h.) %end; ; 

	%let cutn=%eval(&_mgheadern.-2);;
	%do _h=&_mgheadern. %to 3 %by -1; 
		x1=1; y1=(&ymax.-(&_addheadern.-&cutn.)-1)/&ymax.*100; x2=100; y2=y1;
		do i=1 to count(colls&_h.,'#')+1;
			id="myid"||strip(scan(colls&_h.,i,'#'));  output;
		end;
		%let cutn=%eval(&cutn.-1);
	%end; 
run;
data __addcol;
	do __ord=1 to &_addheadern.;
		output;
	end;
run;
data __final;
	set __addcol __final;
run;
%end;

**********************************************************;
%do _i=1 %to &_fnum.;
	%let _coln=&&_fcol&_i.; 
	%let _bandby&_coln.=;
%end;

%if %length(&bandplot.)>0 %then %do;
/*data __tmp_fcol;*/
/*	length attrib value $50.;*/
/*	%do _i=1 %to &_fnum.;*/
/*		%let _coln=&&_fcol&_i.;*/
/*		coln=&_coln.; attrib='LIMITLOWER'; value="&&col&_coln.var1."; output;*/
/*		attrib='LIMITUPPER'; value="&&col&_coln.var3."; output;*/
/*	%end;*/
/*run;*/

data __tmp_bdplt;
length cat attrib $20.;
cat='B'; ord=1; attrib='LIMITLOWER'; output;
cat='B'; ord=2; attrib='LIMITUPPER'; output;
cat='B'; ord=3; attrib='DATATRANSPARENCY'; output;
cat='B'; ord=4; attrib='FILLCOLOR'; output;
cat='B'; ord=5; attrib='FILLTRANSPARENCY'; output;
cat='B'; ord=6; attrib='JUSTIFY'; output;
cat='B'; ord=7; attrib='TYPE'; output;
run;

data __bandplot;
	length bandplot_info $3000. subrefl $1000. subcol $500. attrib value $50.;
	bandplot_info="&bandplot.";
	do n=1 to (count(bandplot_info,'|')+1);
		coln=input(symget(cats('_fcol',n)),best.);
		subrefl=strip(scan(bandplot_info,n,'|'));
		do ord=1 to (count(subrefl,'/')+1);
			subcol=strip(scan(subrefl,ord,'/'));
			attrib=strip(upcase(scan(subcol,1,'=')));
			value=strip(scan(subcol,2,'='));	
			output;
		end;
	end;
run;
proc sql undo_policy=none;
	create table __tmp_bdplt1 as select a.coln,b.cat,b.ord,b.attrib
		from (select distinct coln from __tmp_fcol) as a,__tmp_bdplt as b
		order by a.coln,b.ord;

	create table __bandplot as select distinct a.coln,a.cat,a.ord,a.attrib,b.attrib as attrib1,b.value
		from __tmp_bdplt1 as a left join __bandplot as b
		on a.coln=b.coln and a.attrib=b.attrib
		order by a.coln,a.ord;
quit;

data __bandplot;
	set __bandplot;
	
	if strip(attrib)='LIMITLOWER' then do;
		if value='' then value=strip(symget(cats('col',coln,'var1')));
		call symputx(cats('bdpt',coln,'_LIMITLOWER'),value);
	end;
	if strip(attrib)='LIMITUPPER' then do;
		if value='' then value=strip(symget(cats('col',coln,'var3')));
		call symputx(cats('bdpt',coln,'_LIMITUPPER'),value);
	end;
	if strip(attrib)='DATATRANSPARENCY' then do;
		if value='' then value='0.6';
		call symputx(cats('bdpt',coln,'_datatransparency'),value);
	end;
	if strip(attrib)='FILLCOLOR' then do;
		if value='' then value='blue';
		call symputx(cats('bdpt',coln,'_fillCOLOR'),value);
	end;
	if strip(attrib)='FILLTRANSPARENCY' then do;
		if value='' then value='0.6';
		call symputx(cats('bdpt',coln,'_FILLTRANSPARENCY'),value);
	end;
	if strip(attrib)='JUSTIFY' then do;
		if value='' then value='left';
		call symputx(cats('bdpt',coln,'_JUSTIFY'),value);
	end;
	if strip(attrib)='TYPE' then do;
		if value='' then value='SERIES';
		call symputx(cats('bdpt',coln,'_TYPE'),value);
	end;
run;
data __final;
	set __final;
/*	__ord=_n_-1;*/
	if __ord>=&dropliney. then __ord1=__ord;
/*	call symputx(cats('_colorbands',__ord),cats(Colorbands));*/
run;

*********************************************;

%do _i=1 %to &_fnum.;
%let _coln=&&_fcol&_i.; 
%if %length(%sysfunc(compress(&&bdpt&_coln._LIMITLOWER,,ka)))>0 %then %let _bandby&_coln.=%str(&&_bandby&_coln.#&&bdpt&_coln._LIMITLOWER.);
%if %length(%sysfunc(compress(&&bdpt&_coln._LIMITUPPER,,ka)))>0 %then %let _bandby&_coln.=%str(&&_bandby&_coln.#&&bdpt&_coln._LIMITUPPER.);
%if %length(&&_bandby&_coln.)>0 %then %do;
%if "%substr(%str(&&_bandby&_coln.),1,1)"="#" %then %let _bandby&_coln.=%sysfunc(tranwrd(%substr(%str(&&_bandby&_coln.),2),#,%str(,))); 
/*%put &=_bandby.;*/
proc sql noprint;
	select count(*) into: _bandlimitn from (select distinct &&_bandby&_coln. from __final);
	%if &_bandlimitn.<2 %then %do; 
		%let _bandby&_coln.=; 
		%if %length(%sysfunc(compress(&&bdpt&_coln._LIMITLOWER.,,ka)))>0 %then %do; 
			select &&bdpt&_coln._LIMITLOWER. into: bdpt&_coln._LIMITLOWER from __final where &&bdpt&_coln._LIMITLOWER. >.;
		%end;
		%if %length(%sysfunc(compress(&&bdpt&_coln._LIMITUPPER,,ka)))>0 %then %do; 
			select &&bdpt&_coln._LIMITUPPER. into: bdpt&_coln._LIMITUPPER from __final where &&bdpt&_coln._LIMITUPPER. >.;
		%end;

	%end;
quit;
%end;
%end;

%end;
**********************************************************;

/*%put __________________ before xaxisls ____________;*/

%********************* xaxisls *************************;
%do _i=1 %to &_fnum.;
%let _coln=&&_fcol&_i.; 

%let _logaxis&_coln.=; %let _logaxis&_coln._1=; %let _xmin&_coln.=; %let xmin&_coln.=; %let _xmax&_coln.=; %let xmax&_coln.=;
%let _logaxis_=%scan(&logaxis.,&_i,#|);
%if "&_logaxis_."^="10" and "&_logaxis_."^="2" and "%upcase(&_logaxis_.)"^="E" %then %let _logaxis_=;
%let _logaxis&_coln.=&_logaxis_.;

%if %length(&_logaxis_.)>0 %then %do;
	%if "%upcase(&_logaxis_.)"="2" %then %do; %let _logaxis=%str(log2); %let _logaxis_1=%str(2); %end;
		%else %if "%upcase(&_logaxis_.)"="E" %then %do; %let _logaxis=%str(log2); %let _logaxis_1=%str(2); %end;
		%else %if "%upcase(&_logaxis_.)"="10" %then %do; %let _logaxis=%str(log10); %let _logaxis_1=%str(10); %end;
%end;
%let _xaxisls_=%scan(&xaxisls.,&_i,#|);
%let _xaxislsc_=%scan(&xaxislsc.,&_i,#|);
%let _xaxisls&_coln.=%str(&_xaxisls_.);
%let _xaxislsc&_coln.=%str(&_xaxislsc_.);
%if %length(&_xaxisls_.)>0 %then %do;
	%let xmin&_coln.=%scan(&_xaxisls_.,1,%str( ));
	%let xmax&_coln.=%scan(&_xaxisls_.,%eval(%sysfunc(count(%sysfunc(strip(&_xaxisls_.)),%str( )))+1),%str( ));
%end; %else %do;
************************************************;
proc sql noprint;
	%if %length(&_logaxis_.)>0 %then %do;
	select floor(&_logaxis.(min(&&col&_coln.var1.))) into: _xmin&_coln. from __final where ^missing(&&col&_coln.var1.);
	select ceil(&_logaxis.(max(&&col&_coln.var3.))) into: _xmax&_coln. from __final where ^missing(&&col&_coln.var3.);
	%end; %else %do;
	select floor(min(&&col&_coln.var1.)) into: _xmin&_coln. from __final where ^missing(&&col&_coln.var1.);
	select ceil(max(&&col&_coln.var3.)) into: _xmax&_coln. from __final where ^missing(&&col&_coln.var3.);
	%end;
quit;

data __xaxis;
	length _xaxisls _xaxislsc $500.;
	do i=&&_xmin&_coln.. to &&_xmax&_coln..;
		%if %length(&_logaxis_.)>0 %then %let _numstr=%str(&_logaxis_1.**i);
			%else %let _numstr=%str(i);
		_xaxisls=catx(' ',_xaxisls,&_numstr.);
		_xaxislsc=catx(' ',_xaxislsc,cats('"',&_numstr.,'"'));
	end;
	call symputx("_xaxisls&_coln.",_xaxisls);
	call symputx("_xaxislsc&_coln.",_xaxislsc);
	call symputx("xmin&_coln.",%if %length(&_logaxis_.)>0 %then %do;&_logaxis_1.** %end; &&_xmin&_coln..);
	call symputx("xmax&_coln.",%if %length(&_logaxis_.)>0 %then %do;&_logaxis_1.** %end; &&_xmax&_coln..);
run;
%end;

%end;
/*%put &=logaxis.| &=_logaxis. | &=_logaxis1. | &=_xmin. &=xmin. |&=_xmax. &=xmax. | &=ymax.;*/
%let _sysencoding=%upcase(&sysencoding.);

%do _i=1 %to &_fnum.;
%let _coln=&&_fcol&_i.; 
%if %symexist(col&_coln.title)=0 %then %let col&_coln.title=;
%end;

data __anno_col;
	length id function x1space y1space anchor textcolor textweight textstyle textfont $50. label $300.;
	retain function 'text' x1space "layoutpercent" y1space "datavalue" anchor 'left' textcolor "black";
	set __final;
	textstyle='normal'; textweight='normal'; width=&width.;
	textsize=&fontsize.;  textfont="&textfont."; 
	y1=__ord; 
	if boldind>0 then textweight="bold";
	%do _f=1 %to &columnsn.;
		%if "&&col&_f.type."="T" %then %do;
			id="myid&_f."; x1=&&col&_f.x.; label=strip(vvalue(&&col&_f.var1.)); anchor="&&col&_f.just."; 
			%if &_f.=1 %then %do;
				%if "&_sysencoding."="UTF-8" %then %do;
					if subind=1 then label=unicode('\u2002\u2002')||label;
				%end; %else %if "&_sysencoding."="EUC-CN" %then %do;
					if subind=1 then label=unicode('\u3000\u3000')||label;
				%end;%else %do;
					if subind=1 then label='    '||strip(label);
				%end;
			%end;
			output;
		%end;%else %do;
			if _n_=1 then do;
				id="myid&_f."; x1=&&col&_f.x.; y1=&ymin.; textweight="bold"; label="&&col&_f.title."; output;
			end;
		%end;
	%end;
run;


%if %length(&_colorbands.)>0 %then %do;
%let __col1w=%scan(&columnweights,1,%str( ));
data __final;
	set __final;
	__ord=_n_-1;
	if symexist(cats('_colorbands',__ord)) then Colorbands = input(symget(cats('_colorbands',__ord)),??best.);
	call symputx(cats('_colorbands',__ord),cats(Colorbands));
run;
data anno_POLYGON;
	length id function x1space y1space FILLCOLOR $50.;
    id="myid1"; x1space="WALLPERCENT"; y1space="datavalue"; FILLTRANSPARENCY= 0.6; display="FILL"; 
	%if &_bandsx.=1 %then %do; %let _i_min=0; %let _i_max=&ymax.; %end;
		%else %do; %let _i_min=1; %let _i_max=%eval(&ymax.); %end;
	%do i=&_i_min. %to &_i_max.;
		coln=&i.; type=1;
		order=1; function="POLYGON"; x1=0; y1=&i.-0.5+(&bandoffset.); 
		%if "&_colorbands."="DATA" %then %do;
			if &&_colorbands&i.=1 then do;
				FILLCOLOR="lightgray"; &colorbandsattrs.;
			end; else do; FILLCOLOR="white"; end; output;
		%end;%else %do;
			if mod(&i.,2)=&_bandsx. then do;
				FILLCOLOR="lightgray"; &colorbandsattrs.;
			end; else do; FILLCOLOR="white"; end; output;
		%end;
		order=2; function="POLYCONT"; x1=0; y1=&i.+0.5+(&bandoffset.); output;
		order=3; function="POLYCONT"; x1=100/&__col1w.; y1=&i.+0.5+(&bandoffset.); output;
		order=4; function="POLYCONT"; x1=100/&__col1w.; y1=&i.-0.5+(&bandoffset.); output;
	%end;
	%if %length(&bandplot.)>0 %then %do;
		%do _i=1 %to &_fnum.;
			%let _coln=&&_fcol&_i.; 
			coln=&&_fcol&_i.; type=2;
			%if %length(&&_bandby&_coln.)=0 %then %do;
			id="myid&_coln."; x1space="datavalue"; y1space="datavalue";  display="FILL"; 
			FILLTRANSPARENCY= &&bdpt&_coln._filltransparency.; FILLCOLOR="&&bdpt&_coln._fillcolor."; 
			function="POLYGON"; x1=&&bdpt&_coln._limitlower.; y1=&dropliney.-0.5; output;
			function="POLYCONT"; x1=&&bdpt&_coln._limitlower.; y1=&ymax.+0.5; output;
			function="POLYCONT"; x1=&&bdpt&_coln._limitupper.; y1=&ymax.+0.5; output;
			function="POLYCONT"; x1=&&bdpt&_coln._limitupper.; y1=&dropliney.-0.5; output;
			%end;
		%end;
	%end;
run;
data Anno_polygon1;
	set Anno_polygon(where=(type=1));
	retain colorgr 0;
	if lag(FILLCOLOR)^=FILLCOLOR then colorgr+1;
run;
proc sql undo_policy=none;
	create table Anno_polygon1 as select *,min(coln) as mincoln,max(coln) as maxcoln
		from Anno_polygon1 group by colorgr;
quit;
data Anno_polygon1;
	set Anno_polygon1;
	if (coln=mincoln and order in (1,4)) or (coln=maxcoln and order in (2,3));
run;
proc sort data=Anno_polygon1;
	by colorgr order;
run;
data Anno_polygon;
	set Anno_polygon1 Anno_polygon(where=(type^=1));
run;
%end;
%put --- &_colorbands. ---;
data __anno_;
	set %if %length(&_colorbands.)>0 %then %do;anno_POLYGON %end; __anno_col(where=(label>'')) 
		%if %length(&mergeheader.)>0 %then %do; __anno_col1 anno_mergeheader %end; &add_annods.;
run;
%let yaxislsc=;
%do _x=1 %to &ymax.;
	%let yaxislsc=%str(&yaxislsc. "&_x.");
%end;

%************************************************;
%let _tf=false;
%let _display=none;
%if "&debug."^="0" %then %do;
%let _tf=true;
%let _display=all;
%end;
proc template;
define statgraph ForestPlot; 
	begingraph /axisLineExtent=data;
	layout lattice /columns=&columnsn. columnweights=(&columnweights.) pad=(&pad.) columngutter=&columngutter. BORDER=&_tf.;
		annotate / id="myid"; 
	%do _f=1 %to &columnsn.;
		%if "&&col&_f.type."="T" %then %do;
	    /*-- text column --*/
	    layout overlay /walldisplay=&_display. outerpad=0pt xaxisopts=(display=none) 
			yaxisopts=(linearopts =(viewmin=0 viewmax=&ymax. TickValueSequence=(start=0 end=&ymax. increment=1)) reverse=true display=none);
			annotate / id="myid&_f."; 
	        axistable y=__ord value=__col1 /display=(values);
	    endlayout;
	 	%end;
		%if "&&col&_f.type."="F" %then %do;
	    /*-- figure column--*/
		layout overlay / walldisplay=&_display. outerpad=0pt 
			xaxisopts=(display=(ticks line tickvalues) offsetmin=&_xoffsetmin. offsetmax=&_xoffsetmax.
					%if %length(&&_logaxis&_f.)=0 %then %do;
	                   linearopts =( tickvaluelist=(&&_xaxisls&_f..) %if %length(&&_xaxislsc&_f..)>0 %then %do;
								tickdisplaylist=(&&_xaxislsc&_f..) %end; viewmin =&&xmin&_f.. viewmax=&&xmax&_f..) 
					%end; %else %do;
					   type=log logopts=(base=&&_logaxis&_f. tickvaluelist=(&&_xaxisls&_f..) viewmin =&&xmin&_f.. viewmax=&&xmax&_f..)
					%end;)
	        yaxisopts=(linearopts =(viewmin=0 viewmax=&ymax. TickValueSequence=(start=0 end=&ymax. increment=1)) reverse=true display=none);

			annotate / id="myid&_f."; 
			%if %length(&bandplot.)>0 and %length(&&_bandby&_f.)>0 %then %do;
			bandplot y=__ord1 limitupper=&&bdpt&_f._limitupper. limitlower=&&bdpt&_f._limitlower. /datatransparency=&&bdpt&_f._datatransparency. name="band" 
				display=(fill) justify=&&bdpt&_f._justify. type=&&bdpt&_f._type. fillattrs=(transparency=&&bdpt&_f._filltransparency. color=&&bdpt&_f._fillcolor.); 
			%end;
/*			%put &&col&_f.var1. &&col&_f.var2. &&col&_f.var3. &=dropliney.| &ERRORBARCAPSHAPE.; */
			%let dropliney&_f.=%sysevalf(&dropliney.-&&refline&_f._offset.);
			%put dropliney&_f.=&&dropliney&_f..;
           dropLine x=__refline&_f. y=&&dropliney&_f.. / clip=true datatransparency=&&refline&_f._transparency.
				lineattrs=( color=&&refline&_f._color. thickness=&&refline&_f._thickness. pattern=&&refline&_f._pattern.); 
		   scatterplot x=&&col&_f.var2. y=__ord / subpixel=off markerattrs=(color=&markercolor. symbol=&markersymbol. size=&markersize. 
				%IF %length(&MARKERWEIGHT.)>0 %then %do; WEIGHT=&MARKERWEIGHT. %end;) 
				xerrorupper=&&col&_f.var3. xerrorlower=&&col&_f.var1. ERRORBARCAPSHAPE=&ERRORBARCAPSHAPE. ERRORBARCAPSCALE=&ERRORBARCAPSCALE.
				errorbarattrs=(color=&errorbarcolor. pattern=&errorbarpattern. thickness=&errorbarthickness.) name="scatter";
		endlayout;
		%end;
	%end; 
	endlayout;
	endgraph;
end;
run;
proc sgrender data=__final template=forestplot sganno=__anno_;
run;
%if "&debug."="0" %then %do;
	proc datasets nolist;
		delete __final __anno_: anno_mergeheader: __mergeheader: __tmp: __xaxis __addcol anno_POLYGON __bandplot;
	quit;
%end;
%mend ForestPlot;

*RA Analysis;
*Investigator: ;
*Author: Michelle Marie Wiest;
*2/3/26;
*Data location G:\Shared drives\Biostats\CONFIRM2\Programs\AnalysisDatasets\MPAC2 LOCKED Programs\programs\output\adsl_stand_var;

libname mpac2 "G:\Shared drives\Biostats\CONFIRM2\Programs\AnalysisDatasets\MPAC2 LOCKED Programs\programs\output";
*libname out "C:\Users\MichelleWiest\Documents\GitHub\RA-CONFIRM2"
%include "G:\Shared drives\Biostats\Macros\TableMacro.sas";

*load data;
data dat;
set mpac2.adsl_stand_vars;
run;

proc contents data=dat;
run;
*cool it is "missing" for non-RA patients;

data ra;
set dat;
where pop_mpac2=1;
if pop_RA=1 then RA=1;
if pop_RA ne 1 then RA=2;
run;

*write out demo data in case;
ods excel file="C:\Users\MichelleWiest\Documents\GitHub\RA-CONFIRM2\DemoData.xlsx" options(sheet_name="Top 20 Diffs");
proc print data=ra noobs;
var age bmi dm diabetes dyslipidn ethnicityc famhxcadn racec sexc smoker htnc;
run;
ods excel close;

***********************************************
*Table 1 code starting now...;
***********************************************;
%let INPUT     = ra; 
%let ID 	   = cleerly_id; 
%let STRATA    = RA; 

*** DENOMINATOR COUNT ***;
proc sql noprint;
	select count(distinct(&ID)) into :t  from &INPUT;
	select count(distinct(&ID)) into :t1 from &INPUT where &STRATA=1;
	select count(distinct(&ID)) into :t2 from &INPUT where &STRATA=2;

	select count(distinct(&ID)) format = 5. into :tc   from &INPUT;
	select count(distinct(&ID)) format = 5. into :t1c  from &INPUT where &STRATA=1;
	select count(distinct(&ID)) format = 5. into :t2c  from &INPUT where &STRATA=2; 
quit;

*** PERCENT***;
data _null_;
  t1pct = put(round(&t1/&t, 0.01)*100, 4.1);
  t2pct = put(round(&t2/&t, 0.01)*100, 4.1);
  call symput('t1pct', t1pct);
  call symput('t2pct', t2pct);
run;
 
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(age), where =, id = &ID, sort = 1, label = 'Age (years)', miss=1);
%strata(type =2, dataset  = ra, trt = &STRATA, varname = %str(sexc), where =, id = &ID, sort = 2, label = 'Sex', miss=1);
%strata(type =2, dataset  = ra, trt = &STRATA, varname = %str(ethnicityc), where =, id = &ID, sort = 3, label = 'Ethnicity', miss=1);
%strata(type =2, dataset  = ra, trt = &STRATA, varname = %str(racec), where =, id = &ID, sort = 4, label = 'Race', miss=1);
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(bmi), where =, id=&id, sort=5, label = 'Body Mass Index (BMI), kg/m2', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(smoker), where =, id=&id, sort=7.1, label = 'Smoker', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(dm), where =, id = &ID, sort = 8, label = 'Diabetes or on Diabetes Medication', miss=1);
%strata(type =2, dataset  = ra, trt = &STRATA, varname = %str(htnc), where =, id = &ID, sort = 9, label = 'Hypertension', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(dyslipidn), where =, id = &ID, sort = 10, label = 'Dyslipidemia', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(famhxcadn), where =, id = &ID, sort = 11, label = 'Family history positive for CAD', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(statin_yn), where =, id = &ID, sort = 11, label = 'Statins', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(htn_meds), where =, id = &ID, sort = 11, label = 'Hypertension Medication', miss=1);
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(ldl_base), where =, id = &ID, sort = 11, label = 'LDL at Baseline', miss=1);



*** FORMAT DATA FOR OUTPUT TO RTF ***;
data final;
  length COL_ col1 COL_H $200;
 format  COL_H $200.;
  set   age sexc ethnicityc racec smoker dm htnc htn_meds dyslipidn statin_yn famhxcadn bmi ldl_base;
	if sort0=1 then col_H=col_;
	else col_H="  "||col1;
run;
proc sort data = final; by sort sort0; run;

******************
*** OUTPUT RTF ***
******************;
ODS PATH work.templat(update) sasuser.templat(read)
               sashelp.tmplmst(read);
proc template;
	define style Styles.Custom;
	parent = Styles.Printer;
	replace fonts /
		'TitleFont' = ("Arial",12pt,Bold )
		'TitleFont2' = ("Arial",12pt,Bold Italic)
		'StrongFont' = ("Arial",12pt,Bold)
		'headingFont' = ("Arial",12pt,Bold)
		'docFont' = ("Arial",12pt)
		'footFont' = ("Arial",12pt)
		'FixedStrongFont' = ("Arial",12pt,Bold)
		'FixedHeadingFont' = ("Arial",12pt,Bold)
		'FixedFont' = ("Arial",12pt);
	replace color_list/
	  	'link' = black   		
		'bgH'  = white    		  
		'bgT'  = white			
		'bgD'  = white			
		'fg'   = black			 
		'bg'   = white;			
	replace Table from Output/
		frame = box 			
		rules = all 			
		outputwidth = 75%
		cellpadding = 0pt		
		cellspacing = 0pt 		
		borderwidth = 0.5pt 		
		background  = color_list('bgT');
	end;
  run;



*** USE PROC REPORT TO WRITE THE TABLE TO FILE ***; 

/* 1. Basic Options: Removed 'ls', 'ps', 'formchar' as they don't apply to Excel */
options nonumber nodate missing = " ";

/* 2. Setup ODS Excel with 'Pretty' options */
ods excel file="C:\Users\MichelleWiest\Documents\GitHub\RA-CONFIRM2\Table 1.xlsx"
    /* Pearl is a clean, modern built-in style. You can also try style=Journal for B&W. */
    style=Pearl 
    options(
        sheet_name = "Demographics"  /* Names the tab at the bottom */
        embedded_titles = "yes"      /* Puts the title in Row 1 instead of the print header */
        gridlines = "off"            /* Removes the default gray Excel grid */
        frozen_headers = "yes"       /* Keeps headers visible when scrolling */
        autofilter = "yes"           /* Adds filter arrows to headers */
        row_heights = "20"           /* Sets a comfortable default row height */
    );

ods escapechar='^';

/* 3. Global Title - moved outside to ensure it catches */
title1 justify=left bold font="Arial" height=14pt color="#003366" "Preliminary Demographics"; 

proc report data=final nowd spanrows
    /* General Report Styling */
    style(report)=[frame=void rules=none cellpadding=5pt] /* Removes outer box, adds breathing room */
    split = "|"; 
    
    columns ("Table 1: Prelim Demographics" SORT COL_H COL_1 COL_2); *COL_T; 
    
    /* Sorting definition */
    define sort / order order=internal noprint;
    
    /* Column Definitions */
    /* I converted cellwidth to 'in' (inches) for consistent Excel rendering */
    
    define COL_H / display left "Clinical" 
        style(column)=[font_size=11pt width=3.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt textalign=left color=black borderbottomcolor=black borderbottomwidth=2pt];

    *define COL_T / display center "All|Patients| N = &tc." 
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

    define COL_1 / display center "RA Patients|&t1c.(&t1pct%)"
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

    define COL_2 / display center "Non-RA Patients|&t2c.(&t2pct%)"
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

run;

/* Clear titles and close file */
title;
footnote;
ods excel close;


***************************************************;
* Outcomes Table 2;
**************************************************;
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(TPV), where =, id = &ID, sort = 11, label = 'TPV', miss=1);
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(NCP), where =, id = &ID, sort = 11, label = 'NCP', miss=1);
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(CP), where =, id = &ID, sort = 11, label = 'CP', miss=1);
%strata(type =1, dataset  = ra, trt = &STRATA, varname = %str(LDNCP), where =, id = &ID, sort = 11, label = 'LDNCP', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(present_TPV), where =, id = &ID, sort = 11, label = 'Any Plaque', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(present_NCP), where =, id = &ID, sort = 11, label = 'Any NCP', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(present_CP), where =, id = &ID, sort = 11, label = 'Any CP', miss=1);
%strata(type =3, dataset  = ra, trt = &STRATA, varname = %str(present_LDNCP), where =, id = &ID, sort = 11, label = 'Any LDNCP', miss=1);


*** FORMAT DATA FOR OUTPUT TO RTF ***;
data final;
  length COL_ col1 COL_H $200;
 format  COL_H $200.;
  set   TPV NCP CP LDNCP present_TPV present_NCP present_CP present_LDNCP;
	if sort0=1 then col_H=col_;
	else col_H="  "||col1;
run;
proc sort data = final; by sort sort0; run;

******************
*** OUTPUT RTF ***
******************;
ODS PATH work.templat(update) sasuser.templat(read)
               sashelp.tmplmst(read);
proc template;
	define style Styles.Custom;
	parent = Styles.Printer;
	replace fonts /
		'TitleFont' = ("Arial",12pt,Bold )
		'TitleFont2' = ("Arial",12pt,Bold Italic)
		'StrongFont' = ("Arial",12pt,Bold)
		'headingFont' = ("Arial",12pt,Bold)
		'docFont' = ("Arial",12pt)
		'footFont' = ("Arial",12pt)
		'FixedStrongFont' = ("Arial",12pt,Bold)
		'FixedHeadingFont' = ("Arial",12pt,Bold)
		'FixedFont' = ("Arial",12pt);
	replace color_list/
	  	'link' = black   		
		'bgH'  = white    		  
		'bgT'  = white			
		'bgD'  = white			
		'fg'   = black			 
		'bg'   = white;			
	replace Table from Output/
		frame = box 			
		rules = all 			
		outputwidth = 75%
		cellpadding = 0pt		
		cellspacing = 0pt 		
		borderwidth = 0.5pt 		
		background  = color_list('bgT');
	end;
  run;



*** USE PROC REPORT TO WRITE THE TABLE TO FILE ***; 

/* 1. Basic Options: Removed 'ls', 'ps', 'formchar' as they don't apply to Excel */
options nonumber nodate missing = " ";

/* 2. Setup ODS Excel with 'Pretty' options */
ods excel file="C:\Users\MichelleWiest\Documents\GitHub\RA-CONFIRM2\Table 2.xlsx"
    /* Pearl is a clean, modern built-in style. You can also try style=Journal for B&W. */
    style=Pearl 
    options(
        sheet_name = "Plaque"  /* Names the tab at the bottom */
        embedded_titles = "yes"      /* Puts the title in Row 1 instead of the print header */
        gridlines = "off"            /* Removes the default gray Excel grid */
        frozen_headers = "yes"       /* Keeps headers visible when scrolling */
        autofilter = "yes"           /* Adds filter arrows to headers */
        row_heights = "20"           /* Sets a comfortable default row height */
    );

ods escapechar='^';

/* 3. Global Title - moved outside to ensure it catches */
title1 justify=left bold font="Arial" height=14pt color="#003366" "Plaque Distribution"; 

proc report data=final nowd spanrows
    /* General Report Styling */
    style(report)=[frame=void rules=none cellpadding=5pt] /* Removes outer box, adds breathing room */
    split = "|"; 
    
    columns ("Table 2: Plaque Distribution" SORT COL_H COL_1 COL_2); *COL_T; 
    
    /* Sorting definition */
    define sort / order order=internal noprint;
    
    /* Column Definitions */
    /* I converted cellwidth to 'in' (inches) for consistent Excel rendering */
    
    define COL_H / display left "Plaque Type" 
        style(column)=[font_size=11pt width=3.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt textalign=left color=black borderbottomcolor=black borderbottomwidth=2pt];

    *define COL_T / display center "All|Patients| N = &tc." 
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

    define COL_1 / display center "RA Patients|&t1c.(&t1pct%)"
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

    define COL_2 / display center "Non-RA Patients|&t2c.(&t2pct%)"
        style(column)=[font_size=11pt width=1.5in asis=on vjust=middle] 
        style(header)=[background=#f0f0f0 font_weight=bold font_size=11pt color=black borderbottomcolor=black borderbottomwidth=2pt];

run;

/* Clear titles and close file */
title;
footnote;
ods excel close;

























*old code look weird;

options nonumber nodate ls=200 ps=80 missing = " "
   formchar="|----|+|---+=|-/\<>*" orientation=landscape;
ods excel file= "C:\Users\MichelleWiest\Documents\GitHub\RA-CONFIRM2\Table 1.xlsx" style=custom ;
ods escapechar='^';

 proc report  data = final   nowindows    spacing=1    headline    headskip    split = "|"   ;
    columns ( "Table 1: Prelim Demographics" SORT COL_H COL_T COL_1 COL_2);
  define sort    	/order order = internal noprint;
	   define COL_H    	/display left  "Clinical" 
						style(column)=[ font_size=12pt cellwidth=65 asis=on] 
						style(Header)=[ cellwidth=65 font_weight=bold font_size=12pt  asis=on];
	   define COL_T     /display center "All|Patients| N = &tc." 
						style(column)=[ font_size=12pt cellwidth=32  asis=on] 
						style(Header)=[ cellwidth=32 font_weight=bold font_size=12pt asis=on];
	   define COL_1     /display center "Patients|US|&t1c.(&t1pct%)" 
						style(column)=[ font_size=12pt cellwidth=32  asis=on] 
						style(Header)=[ cellwidth=32 font_weight=bold font_size=12pt asis=on];
	   define COL_2     /display center "Patients|OUS|&t2c.(&t2pct%)" 
						style(column)=[ font_size=12pt cellwidth=32 asis=on] 
						style(Header)=[ cellwidth=32 font_weight=bold font_size=12pt asis=on];
     title1 justify=center bold font=ArialBlack height=14pt  "CONFIRM2"; 
  run; 
 

title;
footnote;
 ods excel close;  




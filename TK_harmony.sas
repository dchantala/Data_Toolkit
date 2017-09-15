/*--------------------------------------------------------------------------------------------------------------------------------------
* program	tk_harmony.sas                                                                         
* function	SAS macro measures the harmmony  of two SAS data sets and reports differences in length, type, and labels 
*		     of variables common to both data sets, and variables unique to either of the sets.                                                              
* author	     Kim Chantala                          
*             Copyright RTI International 2017
*             RTI is a registered trademark and trade name of Research Triangle Institute
* version     v2                                                   
* usage:                                                                                 
*		%tk_harmony(set1, set1_id, set2, set2_id);                      
*		where:                                                             
*                      set1 = name of 1st data set;                                               
*                      set1_id = Identifying abbreviation (<4 characters) for set 1 used in output 
*                      set2 = name of 2nd data set                                               
*                      set2_id = Identifying abbreviation (<4 characters) for set 2 used in output 
*                      out = name of output data set with harmony measures.
*------------------------------------------------------------------------------------------------------------------------------------------- */

%macro tk_harmony(set1, set1_id, set2, set2_id, out);

proc format; value vartype 1='NUM' 2='CHAR' .=' '; run;

* Step 1) Create data set with measures of harmony;

proc contents data=&set1 noprint out=info1(keep=label length name type varnum); run;
proc contents data=&set2 noprint out=info2(keep=label length name type varnum); run;

data info1; set info1; name=upcase(name); run;
data info2; set info2; name=upcase(name); run;
proc sort data=info1; by name; run;
proc sort data=info2; by name; run;

data _join_;

length harmony $ 10;
length type_match $ 3;
length label_match $ 3;
length location $20 type_match $3 label_match $3;

merge info1(in=in1 rename=(label=label1 length=length1 type=type1 varnum=varnum1))
	info2(in=in2 rename=(label=label2 length=length2  type=type2 varnum=varnum2));
by name;

length type_length1 $9 type_length2 $9;
if not missing(type1) and not missing(length1) then type_length1= compress(put(type1,vartype.))||' '||compress(length1); else type_length1=' ';
if not missing(type2) and not missing(length2) then type_length2= compress(put(type2,vartype.))||' '||compress(length2); else type_length2=' ';

if in1 and in2 then location="Both";
else if in1 and not in2 then location="&set1_id ";
else if in2 and not in1 then location="&set2_id ";

if in1 and in2 then do;
	if compress(upcase(label1))=compress(upcase(label2)) then label_match='Yes'; else label_match='No';
	if type1=type2 and length1=length2 then type_match='Yes'; else type_match='No';
end;
else do;
	label_match = 'N/A';
	type_match='N/A';
end;

if in1 and in2 and type_match='Yes' and label_match='Yes' then harmony='SAME'; 
else if in1 and in2 and (type_match='No' or label_match='No') then harmony='DIFF';
else if in1 and not in2 then harmony='SOLO';
else if in2 and not in1 then harmony='SOLO';

label location="Variable location";
label harmony="Harmony measure comparing type, length, and label of variable (DIFF = different, SAME = All Match, SOLO = variable in only one file";
label label_match="Both have same variable label?";
label type_match="Both have same data type?";
label type_length1="Data type for &set1_id";
label type_length2="Data type for &set2_id";
label label1 = "Label for &set1_id";
label label2 = "Label for &set2_id";
run;

*Step 2) Write Report;

proc tabulate data=_join_ missing;
class harmony location type_match label_match;
table harmony*location*type_match*label_match all, all*(n*f=8. reppctn)/rts=10 box="Harmony of Variables: &set1_id=&set1 and &set2_id=&set2" ;
keylabel RepPctN = 'Percent';
keylabel all = 'Total';
label harmony = 'Harmony measure';
run;

/*ods startpage=now;*/

proc report data=_join_ headline wrap split="~";
where harmony in ('SOLO', 'DIFF');
column (" &set1_id=&set1 ~ &set2_id=&set2 ~ ~ Inharmonic Variables" harmony name location) 
	('Data Type' type_match type_length1 type_length2) 
	('Label' label_match label1 label2);
define harmony / display "Harmony Measure" center;
define name / display "Variable name" center;
define location / display "Variable Location" center; 
define type_match / display "Same?" center;
define type_length1 / display "&set1_id" center  style(column)={tagattr="format:@"};
define type_length2 / display "&set2_id" center  style(column)={tagattr="format:@"};
define label_match / display 'Same?' center;
define label1 / display "&set1_id" center;
define label2 /display "&set2_id" center;
compute harmony;
	if harmony in ("DIFF", "SOLO") then call define(_ROW_, "style", "style=[fontweight=bold]");
endcomp;
run;

* Step 3) Create output data set if requested;
%if "&out" ^="" %then %do;
	data &out (label="Harmony measure for &set1 (&set1_id) vs &set2 (&set2_id)");
	set _join_;
	label location="Variable location";
	label harmony="Harmony measure comparing type, length, and label of variable (DIFF = different, SAME = All Match, SOLO = variable in only one file)";
	label name = "Variable Name";
	label label_match="Both have same variable label?";
	label type_match="Both have same data type?";
	label type_length1="&set1_id variable data type (Num, Char) and length";
	label type_length2="Data type for &set2_id variable data type (Num, Char) and length";
	label label1 = "&set1_id variable label";
	label label2 = "&set2_id variable label";

	keep location harmony name label_match type_match type_length1 type_length2 label1 label2;
	run;
%end;

ods listing;

%mend tk_harmony;




/*--------------------------------------------------------------------------------------------------------------------------------------
* program	tk_dataset_list.sas                                                                        
* function	Create a list of data sets & attributes for files in a folder                                                              
* author      	Kim Chantala                          
*             	Copyright RTI International 2017 
*             	RTI is a registered trademark and trade name of Research Triangle Institute
* version 	2017.09.15                                                        
* usage                                                                                 
*		%tk_dataset_list(libref);                      
*		where:                                                             
*                      libref = name associated with the folder pathname by the LIBNAME statement ;                                               
*--------------------------------------------------------------------------------------------------------------------------------------- */

%macro tk_dataset_list(libref);

proc datasets lib=&libref noprint;

contents data=_all_ out=work._ds_theInfo;
run;


/*proc contents data=work._ds_theInfo;
run;

proc freq data=work._ds_theInfo;
tables memname/list missing;
tables sorted sortedby/list missing;
run;

proc means data=work._ds_theInfo;
class memname;
run;

proc print data=work._ds_theInfo (obs=30);
var name sorted sortedby nodupkey noduprec;
run;
*/

proc freq data=work._ds_theInfo noprint;
tables memname*memlabel*crdate*nobs/list missing out=count(drop=percent rename=(count=NVARS memname=SetName));
run;

data _ds_thefiles;
set count;
setname=upcase(setname);
label NVARS='Number of Variables';
label SetName='SAS Data set name (*.sas7bdat)';
run;

proc sort data=_ds_thefiles; by setname; run;

proc report data=_ds_thefiles nocenter nowindows headline wrap;
column  setname memlabel crdate nobs nvars;
define memlabel/display  width=50;
define crdate/display center width=15;
define nobs/display center width=10 missing;
define nvars/display center width=10 missing;
format nobs comma15.0;

run;

proc datasets noprint;
delete _ds_thefiles _ds_theInfo;
run;
%mend dataset_list;

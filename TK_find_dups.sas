/*--------------------------------------------------------------------------------------------------------------------------------------
* program	tk_find_dups.sas                                                                         
* function	Find observations with duplicate ID values.                                                                 
* author	Kim Chantala                          
*             	Copyright RTI International 2017
*             	RTI is a registered trademark and trade name of Research Triangle Institute
* version       2017.09.15                                                  
* usage                                                                                 
*		%tk_find_dups(dataset, one_rec_per, dup_output)                    
*		where:                                                             
*               Required Variables:                                                                                                                 
*                    dataset= Name of Data set being examine for duplicate values of Identification Variables                          
*                    one_rec_per = Variable specification for variables that identify unique observations.                             
*                                  For example, if the data set should have only one record for every person (CASEID) for every        
*                                  wave of data (WAVE) then one_rec_per = CASEID*WAVE.                                                  
*              Optional Variables:                                                                                                     
*                    dup_output = name of data set to return a copy of the duplicate values of one_rec_per  
*------------------------------------------------------------------------------------------------------------------------------------------- */

%macro tk_find_dups(dataset, one_rec_per, dup_output);

proc freq data=&dataset noprint;
tables &one_rec_per/list missing out=_dup_check_(rename=(count=COPIES));
format _all_;
run;

%let dsid=%sysfunc(open(&dataset));                                                                                         
%if &dsid %then %do;
	%let numobs=%sysfunc(attrn(&dsid,nlobsf));                                                                                                 
	%let rc=%sysfunc(close(&dsid));                                                                                                         
%end;

proc freq data=_dup_check_;
title5 "Data set being examined:  &dataset (N=&numobs)";
title6 "Identification variables:  &one_rec_per";
title7 "There should be only one record for every unique value of &one_rec_per";
tables copies/list missing;
label copies="COPIES = Number of records with identical values of &one_rec_per";
format _all_;
run;

data _the_dups_;
set _dup_check_;
if copies>1;
run;

/* If there are no duplicates, print message and return (no Duplicates to print)*/
%let dsid=%sysfunc(open(_the_dups_));                                                                                         
%if &dsid %then %do;
	%let numobs=%sysfunc(attrn(&dsid,nlobsf));                                                                                                 
	%let rc=%sysfunc(close(&dsid));                                                                                                         
%end;
%if &numobs=0 %then %do;
%put NOTE: No duplicates, Only one record per &one_rec_per in &dataset data set.  ;
%return;
%end;

/* Duplicates found, print and return data set if requested*/

%put WARN: Duplicates found, Multiple records have same value of &one_rec_per in &dataset data set.  ;

proc freq data=_the_dups_;
title5 "Values of &one_rec_per occurring on more than one record in data set &dataset";
title6 "COPIES = Number of records with identical values of &one_rec_per";
tables &one_rec_per*copies /list missing nocum nofreq nopercent nocol;
format _all_;
run;
title3 ' ';

%if &dup_output ^= "" %then %do;
	data &dup_output (label="Values of &one_rec_per for duplicate observations in &dataset" drop=percent);
	set _the_dups_;
        label copies = "Number of records with identical values of &one_rec_per";
	run;
%end;
%mend;



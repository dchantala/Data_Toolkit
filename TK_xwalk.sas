/*--------------------------------------------------------------------------------------------------------------------------------------
* program	tk_xwalk.sas                                                                          
* function	Create a crosswalk showing the variables in a group of files                                                              
* author	Kim Chantala                          
*             	Copyright RTI International 2017
*             	RTI is a registered trademark and trade name of Research Triangle Institute
* version       2017.09.15                                                
* usage:                                                                                 
*		%tk_xwalk(SetList);                      
*		where:                                                             
*                      setList = name of 1st data set;                                               
*------------------------------------------------------------------------------------------------------------------------------------------- */

%macro tk_xwalk(SetList);

%local i next_name;

%do i=1 %to %sysfunc(countw("&SetList"," "));   *&num_list;
	%let next_name = %scan(&SetList, &i, ' ');
	proc contents data=&next_name noprint out=info(keep=nobs libname memname name label type length format) ;
	data _xw_thevars_;
	%if &i = 1 %then %do; 
		set info(in=in_info); 
	%end;
	%else %do;
		set _xw_thevars_ info(in=in_info);
	%end;
	label memname='Data Set Name (*.sas7bdat)';
	length vartype $ 10;
	length which_set $ 32;
	if type=1 then vartype=compress("NUM-"||length);
	else if type=2 then vartype=compress("CHAR-"||length);
	label vartype='Type-Length';
	if in_info then do;
		number=&i;
		which_set = "&next_name";
		label which_set='Data Set';
	end;
	run;
%end;

data _xw_thevars_; 
set _xw_thevars_;
name=upcase(name);
run;

proc tabulate data=_xw_thevars_(keep= nobs memname which_set name label type length format vartype) missing;
class _all_;
tables all name*label*vartype*format , (all which_set)/rts=15 box='VARIABLE CROSSWALK';;

keylabel n='Frequency';
keylabel all='Total' ;
run;

proc datasets noprint;
delete _xw_thevars_;
run;
%mend;


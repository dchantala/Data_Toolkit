/*--------------------------------------------------------------------------------------------------------------------------------------
* program	tk_codebook.sas                                                                        
* function	Create a CODE BOOK for SAS data sets.                                                             
* author	    Kim Chantala       
* 			Copyright 2017 RTI International
* Note		An earlier version of tk_codebook.sas was released under the name of proc_codebook.sas (2010, Jim Terry & Kim Chantala, UNC Chapel Hill NC).;
*                 
* version	2017.10.16                                                     
* usage:                                                                                 
*		%tk_codebook(
			lib=SAS_data,
			file1= ,
			fmtlib=work,
			cb_type=XLSX,
			cb_file=&cb_name,
			var_order=internal,
    		cb_output = my_codebook,
			cb_size=BRIEF,
			organization = One record per CASEID,
			include_warn=YES);                                                   
      where: 
	LIB = name of library for SAS data set (see FILE1 variable) used to create the codebook
	FILE1 = name of SAS data set used to create the codebook
	FMTLIB = 2-level name of format library
	CB_TYPE = type of codebook (XML, XLSX, XLS, PDF, RTF)
	CB_FILE = name of file for the codebook being created
	Optional Variables: 
	CB_OUTPUT = Name of data set used to return a copy of the data used to create the FULL codebook.  
	VAR_ORDER = ordering of variables in codebook: CUSTOM_ORDER (order from file named work.custom_order), 
	INTERNAL (order of variables in data set), ommitted = alphabetical as determined by PROC REPORT. 
	ORGANIZATION = text indicating the organization of observations in the data set.  For example,
	there might be one record per CASEID or one record per CASEID*WAVE.
	INCLUDE_WARN= flag to control printing of WARNING messages in Codebook (in addition to LOG file)
	YES=prints warnings in file specified by CB_FILE (default), 
	OTHER=warnings printed only in LOG file. 
	CB_SIZE = indicate reduced size (CB_SIZE=BRIEF) or full size (CB_SIZE) of codebook. A BRIEF codebook has 
	columns of information in the printed codebook.

** TITLES AND FOOTNOTES
	User Specified:  TITLE1, TITLE2 and all FOOTNOTES are specified by user.
	DTK_CODEBOOK Specified: 
	TITLE4 lists the number of observations in Data set.
	TITLE5 lists the number of variables in the data set. 
	TITLE6 lists the organization of the data set, with the ORGANIZATION specified in the 
	global macro variable ORGANIZATION;

 ** RECOMMENDED REQUIREMENTS FOR THE SAS FILE TO BE DOCUMENTED:
	1. Have labels on all variables
	2. Have user FORMATs assigned to all categorical variables
	3. Have a data set label on the SAS file
	4. By default, the codebook is ordered by Variable Name.  See the instructions in ORDERING VARIABLES to control
	order of variables are listed in codebook.

** ORDERING VARIABLES IN CODEBOOK
	To choose an order other than alphabetical or internal order, create a simple two variable file called
	work.custom_order before you call the macro.  The first variable is NAME, a 32 
	character field with your variable name in UPPER CASE. The second variable is ORDER, a 
	numeric field with the order you want the  variables to print, 1 to N.  see the following example:

		data order;
		length name $ 32;
		name = "T1    "; ORDER = 1; OUTPUT;
		name = "HHID09"; ORDER = 2; OUTPUT;
		name = "LINE09"; ORDER = 3; OUTPUT;
		name = "H1D   "; ORDER = 4; OUTPUT;
		run;
*--------------------------------------------------------------------------------------------------------------------------------------- */


%global cb_type;
%global cb_file;
%macro tk_codebook(lib,file1,fmtlib, var_order, cb_type, cb_file, cb_output, organization, include_warn=YES, cb_size=FULL);
	ods listing close;

	/** SECTION 1 - Get the PROC CONTENTS data**/
	proc contents data=&lib..&file1 noprint 
		out=_cb_var_info_(keep=name format formatl formatd type length nobs label memlabel crdate varnum);
	run;

	data _cb_var_info_;
		set _cb_var_info_;
		length temp $ 16;
		name = strip(upcase(name));

		if formatl > 0 and formatd ge 0 and formatd le 9 then
			do;
				substr(temp,1,2) = put(formatl,2.0);
				substr(temp,3,1) = ".";
				substr(temp,4,1) = put(formatd,1.0);
				format = trim(format)||temp;
			end;

		drop temp;

	proc sort data=_cb_var_info_ out=_cb_file_formats_ (keep=format) nodupkey;
		by format;
	run;

	proc sort data=_cb_var_info_ (rename=(type=ntype));
		by name;

		/** SECTION 2 - Get the PROC MEANS Data, N, MEAN, STD_DEV, MIN, and MAX **/
		ods output 'Summary Statistics'=_cb_means_;

		/*** Execute PROC MEANS on the Longitudinal File, it produces 1 long observation  ***/
	proc means data=&lib..&file1;
	run;

	ods output close;
	run;

	/*** Rewrite the PROC MEANS o/p file as 1 observation per variable ***/
	data _cb_mean_results_;
		set _cb_means_;
		length name $ 32 desc $ 50;

		/** ARRAY for Numeric Variables, N, MEAN, STD_DEV, MIN, and MAX **/
		array x {*} _numeric_;

		/** ARRAY for Character Variables, VARIABLE and LABEL **/
		array z {*} _character_;

		/** Determine Number of Character and Numeric Variables **/
		cv = dim(z);
		nv = dim(x);
		vars = (dim(z) / 2) - 1; /* Substract Out the 2 Computed Variables, variable and label*/

		/** O/P 1 observation per variable per by group **/
		do i = 1 to vars;
			if i = 1 then
				do;
					numeric = 1;
					alpha = 1;
				end;

			n = x(numeric);
			mean = x(numeric+1);
			std_dev = x(numeric+2);
			min = x(numeric+3);
			max = x(numeric+4);
			name = z(alpha);
			desc = z(alpha+1);
			name=upcase(name);
			output;
			numeric = numeric + 5;
			alpha   = alpha + 2;
		end;

		keep name desc n mean std_dev min max;
	run;

	proc sort data=_cb_mean_results_;
		by name;
	run;

	/** SECTION 3 - Merge the PROC MEANS and PROC CONTENTS results **/
	data _cb_var_merge_ (drop=desc);
		merge _cb_var_info_ (in=a) _cb_mean_results_ (in=b);
		by name;
		length range $ 40 c_max c_min $ 12;

		if max - int(max) > 0 then
			max = round(max,.01);

		if min - int(min) > 0 then
			min = round(min,.01);
		c_max = put(max,11.2);
		c_min = put(min,11.2);
		q = index(c_min,".");

		if q = 0 then
			q = 1;

		if substr(c_min,q,2) = ".0 " or substr(c_min,q,2) = ".00" then
			substr(c_min,q,3) = "   ";
		q = index(c_max,".");

		if q = 0 then
			q = 1;

		if substr(c_max,q,2) = ".0 " or substr(c_max,q,2) = ".00" then
			substr(c_max,q,3) = "   ";
		range = compress(c_min||' to '||c_max);

		/*q = index(range,"-"); * put "Range Length " n=;*/
		if q > 7 then
			range = trim(left(c_min))||' to '||trim(left(c_max));

		if ntype = 1 then
			type = 'Num ';

		if ntype = 2 then
			type = 'Char';
		drop ntype c_max c_min q;
	run;

	/** SECTION 4 - Create macro variables with number of obs and variables for the 
	              Codebook titles and make all variable names Upper Case, also
	              create macro variables with the data set label and dataset
	              creation date/time for use in the title;**/
	data _cb_var_data_ (drop=nvars nobs memlabel crdate);
		set _cb_var_merge_ end=lastone;
		retain nvars;

		if _n_ = 1 then
			do;
				nvars = 0;

				/* Following 3 items from PROC CONTENTS*/
				call symputx('obs_cnt',put(nobs,8.0));
				call symputx('ds_label',memlabel);
				call symputx('ds_date',put(crdate,datetime16.));
			end;

		nvars+1;
		name = compress(upcase(name));

		if lastone = 1 then
			call symputx('var_cnt',put(nvars,6.0));
	run;

	/** SECTION 5 - Create a macro variable list of all the variables that
	              have a FORMAT (Frequencies will only be generated for
	              numeric variables that have a format) and run PROC FREQ
	              for those variables ; **/
	proc format library=&fmtlib 
		cntlout=_cb_fmt_defs_(keep=fmtname type start end label hlo eexcl sexcl 
		rename=(fmtname=format 
		type=fmt_type
		start=fmt_start
		end=fmt_end
		label=desc
		hlo=fmt_hlo
		eexcl=fmt_end_excl
		sexcl=fmt_start_excl));
	run;

	data _cb_fmt_defs_2_;
		set _cb_fmt_defs_;
		length fmt_hold $ 33;
		format=compress(upcase(format));

		if fmt_type =: 'C' then
			do;
				fmt_hold = format;
				substr(format,1,1) = '$';
				substr(format,2,30) = fmt_hold;
			end;

		seq = _n_;

	proc sort data=_cb_fmt_defs_2_;
		by format seq;

	data _cb_fmt_defs_ (drop=seq);
		merge _cb_file_formats_ (in=a) _cb_fmt_defs_2_ (in=b);
		by format;

		/** Only KEEP Formats that appear in the file AND are in the USER FORMAT LIBRARY;**/
		if a = 1 and b = 1 then;
		else delete;

	/*** Create 2 data sets, one with format definitions and one with format ranges;**/
	data _cb_fmt_defs_ (drop=range frange) _cb_fmt_labels_ (keep=format desc frange) _cb_test_fmt_labels_;
		set _cb_fmt_defs_;
		length range frange $ 40;  /* 40 is max size of format value text*/

		if compress(fmt_start) = compress(fmt_end) then
			range = trim(left(fmt_start));

		if compress(fmt_start) ne compress(fmt_end) then
			do;
				range = (trim(left(fmt_start))||'-'||trim(left(fmt_end)));
			end;

		frange = range;
	run;

	proc contents data=_cb_fmt_labels_ noprint out=_cb_fmt_val_length_(keep=name length);
	run;

	data _null_;
		set _cb_fmt_val_length_;

		if name='desc';
		call symputx('n_value_lab',put(length,8.0));
	run;

	/*** Use format definitions, _cb_fmt_defs_, to identify categorical variables (NUM & CHAR) to pass to PROC FREQ ***/
	/* Sort the user formats from the user format library.*/
	proc sort data=_cb_fmt_defs_  out=_cb_user_formats_ (keep=format) nodupkey;
		by format;

		/*Sort the variable names by format, some variables will not have formats.*/
	proc sort data=_cb_var_data_ (keep=name format type length label) out=_cb_formatted_vars_;
		by format;

		/* Keep variables that have a user format and variables that do not in separate data sets.*/
	data _cb_user_formatted_variables_ (drop=type length label) _cb_char_var_no_fmt _cb_num_var_no_fmt;
		merge _cb_user_formats_ (in=a) _cb_formatted_vars_ (in=b);
		by format;

		if a=1 then
			output _cb_user_formatted_variables_;

		if type='Char' and a=0 then
			output _cb_char_var_no_fmt;
		else if type='Num' and a=0 then
			output _cb_num_var_no_fmt;
	run;

	/*Create a macro variable that is a list of all variables with user formats.*/
	proc sql noprint;
		select name into :tlist separated by ' ' from _cb_user_formatted_variables_ where not missing(format);
	quit;

	/* Check to see if there is at least one variable with an assigned format - quit processing if no assigned format*/
	%let dsid=%sysfunc(open(_cb_user_formatted_variables_));

	%if &dsid %then
		%do;
			%let numobs=%sysfunc(attrn(&dsid,nlobsf));
			%let rc=%sysfunc(close(&dsid));
		%end;

	%if &numobs=0 %then
		%do;
			%put ERROR: Codebook not created. No User Assigned FORMATS. Assign user defined format to one or more variables in "&file1" data set.;

			%return;
		%end;

	ods output 'One-Way Frequencies'=_cb_freqs_;

	proc freq data=&lib..&file1;
		table &tlist / missing;
	run;

	ods output close;
	run;

        /** If the data set contained variables with the names of FREQUENCY, PERCENT, CUMFREQUENCY, CUMPERCENT then PROC FREQ              **/
        /** placed the statistic having the same name in variables named FREQUENCY2, PERCENT2, CUMFREQUENCY2, CUMPERCENT2 in the           **/
        /** output data set _cb_freqs_.  For all other variables, these statistics were placed in variables named  FREQUENCY, PERCENT,     **/
        /** CUMFREQUENCY, CUMPERCENT.   The next data set fixes this so TK_codebook.sas will have the statistics in the expected variable  **/
        /** in _cb_freqs_.                                                                                                                 **/

	data _cb_freqs_;
	set _cb_freqs_;
/*         %IF %UPCASE(%SCAN(table, 2, ' ')) EQ "PERCENT"  %THEN
         %DO;
           IF UPCASE(SCAN(table, 2, ' ')) EQ "PERCENT"  AND ~MISSING(percent2) THEN percent = percent2 ;
         %END ;

         %IF %UPCASE(%SCAN(table, 2, ' ')) EQ "FREQUENCY" %THEN
         %DO ;
           IF UPCASE(SCAN(table, 2, ' ')) EQ "FREQUENCY" AND ~MISSING(frequency2) THEN frequency = frequency2 ;
         %END ;

         %IF %UPCASE(%SCAN(table, 2, ' ')) EQ "CUMFREQUENCY" %THEN
         %DO ;
           IF UPCASE(SCAN(table, 2, ' ')) EQ "CUMFREQUENCY" AND ~MISSING(cumfrequency2) THEN cumfrequency = cumfrequency2 ;
         %END ;

         %IF %UPCASE(%SCAN(table,2,' '))  EQ  "CUMPERCENT" %THEN
         %DO ;
           IF UPCASE(SCAN(table,2,' '))  EQ  "CUMPERCENT" AND ~MISSING(cumpercent2) THEN cumpercent = cumpercent2 ;
         %END ;
*/
	if upcase(scan(table,2,' ')) = "PERCENT" and not missing(percent2) then percent=percent2;
	else if upcase(scan(table,2,' ')) = "FREQUENCY" and not missing(frequency2) then frequency=frequency2;
	else if upcase(scan(table,2,' ')) = "CUMFREQUENCY" and not missing(cumfrequency2) then cumfrequency=cumfrequency2;
	if upcase(scan(table,2,' ')) = "CUMPERCENT" and not missing(cumpercent2) then cumpercent=cumpercent2;

	format frequency cumfrequency best8.;
	format percent cumpercent 6.2 ;
	run;


	/** SECTION 6 - Create a variable name based frequency file for 
	              merging with the MEANS and CONTENTS data. **/
	data _cb_freq_results_;
		set _cb_freqs_;

		*length desc $ 80 _tmpname_ $ 32;
		length desc $ &n_value_lab  _tmpname_ $ 32;

		/** ARRAY for Character Variables, VARIABLE and LABEL **/
		array z {*} _character_;
		cv = dim(z);
		missed = 0;

		do i = 1 to cv;
			if not missing(z(i)) then
				missed+1;
		end;

		desc = '                                    ';

		do i = 2 to cv;	/* The 1st Variable is the Table variable. */
			if not missing(z(i)) then
				do;
					desc = z(i);
					i = cv;
				end;
		end;

		cnt = _n_;   /* This variable is used to keep everything in sort order. */

		/* Extract the variable Name from the Table Variable. */
		substr(table,1,5) = '     ';

		table = left(table);
		_tmpname_ = upcase(table);
		keep _tmpname_ desc frequency percent CumFrequency CumPercent cnt;
	run;

	/** SECTION 7 - Merge the Frequencies with the MEANS and CONTENTS data. **/
	proc sort data=_cb_freq_results_;
		by _tmpname_ cnt;

	proc sort data=_cb_var_data_;
		by name;

	data _cb_fmt_merge_(rename=(_tmpname_=name));
		merge _cb_var_data_ (in=a rename=(name=_tmpname_)) _cb_freq_results_ (in=b );
		by _tmpname_;

		if b=1 then
			has_freq_results=1;
		else has_freq_results=0;
	run;

	** SECTION 8 - Merge the Format Ranges, based on format and format label (desc),
	              note Format Ranges will override the Min/Max range;
	proc sort data=_cb_fmt_merge_;
		by format desc;

	proc sort data=_cb_fmt_labels_;
		by format desc;

		%max_lengths(set1=_cb_fmt_labels_, set2=_cb_fmt_merge_);

	data _cb_ranges_;
		length &char_lengths;
		merge _cb_fmt_labels_ (in=a) _cb_fmt_merge_ (in=b);
		by format desc;

		if a=1 then in_labels=1;
		else in_labels=0;

		if b=1 then in_merge=1;
		else in_merge=0;

		if a = 1 and b = 1 and has_freq_results=1 then
			do;
				range = frange;
			end;
		else if a ne 1 and b eq 1 and has_freq_results=1 then
			do;
				range=frange;

				if compress(desc) in ('.') then
					desc='SAS missing (.)';
				else if compress(desc) in (' ') then
					desc='Missing (blank)';
			end;
		else if has_freq_results=0 and (compress(desc)) in (' ') and compress(range) ne '-' and compress(type)='Num'  then
			do;
				desc='Range';
				CumFrequency=&obs_cnt;
				CumPercent = 100.00;
				/*CumPercent = round(((n / &obs_cnt) * 100),.01);*/
				cnt=1;
			end;
		else if has_freq_results=0 and compress(type)='Char' then
			do;
				desc='Blank, Text, or Value supplied';
				Frequency=&obs_cnt;
				percent = 100.00;
				cumFrequency=&obs_cnt;
				cumPercent=100.00;
				cnt=1;
				range = '**OTHER** ';
			end;

		created=0;
		output;

		if has_freq_results=0 and type='Num' then
			do;
				/*Create record for missing values when results come from proc means. */
				Frequency=&obs_cnt-n;
				percent  = round(((frequency / &obs_cnt) * 100),.01);
				n = Frequency;
				range = ' ';
				cumFrequency=n;
				cumpercent=percent;
				desc='SAS missing (.)';
				cnt=0;
				created=1;
				output;
			end;

		drop frange;
	run;

	proc sort data=_cb_ranges_;
		by name format cnt;

		/** SECTION 9 - Initial prep for printing. **/
		options nocenter mergenoby=error linesize=170 pagesize=40 ORIENTATION=LANDSCAPE nofmterr missing=" ";

	data _cb_all_data_ (drop=warning) 
			_cb_msgs1_ (keep=warning name label format desc frequency percent cumfrequency cumpercent)
			_cb_msgs2_ (keep=warning name label format desc frequency percent cumfrequency cumpercent)
			_cb_msgs3_ (keep=warning name label format desc min max frequency percent cumfrequency cumpercent)
			_cb_msgs4_ (keep=warning name label format desc frequency percent cumfrequency cumpercent);
		set _cb_ranges_  end=lastone;
		length WARNING $ 100;
		length mean_char $ 15;
		length type_length $9;

		/*range=strip(range);
		*desc=strip(desc);*/
		type_length= compress(type)||' '||compress(length);

		if format =: 'MMDDYY' and has_freq_results ne 1 then
			mean_char=put(int(mean),mmddyy10.);
		else if format=: 'DATETIME'  and has_freq_results ne 1  then
			mean_char=put(datepart(int(mean)), mmddyy10.);
		else if format =: 'TIME' and has_freq_results ne 1  then
			mean_char=put(int(mean),time5.);
		else if .<max<10 then
			mean_char=put(mean,f5.2);
		else if .<max<100 then
			mean_char=put(mean,f6.2);
		else if .<max<1000 then
			mean_char=put(mean,f7.2);
		else if .<max<10000 then
			mean_char=put(mean,f8.2);
		else mean_char=put(mean,best10.);
		label mean_char='Mean';

		if format =: 'DATETIME'  and has_freq_results ne 1 and (min ne .) and (max ne .) then
			do;
				range = '                           ';
				substr(range,1,10) = put(datepart(min),mmddyy10.);
				substr(range,11,4) = " to ";
				substr(range,15,10) = put(datepart(max),mmddyy10.);
			end;

		if format =: 'MMDDYY'  and has_freq_results ne 1 then
			do;
				range = '                           ';
				substr(range,1,10) = put(min,mmddyy10.);
				substr(range,11,4) = " to ";
				substr(range,15,10) = put(max,mmddyy10.);
			end;

		if format =: 'TIME' and has_freq_results ne 1 then
			do;
				range = ' ';
				substr(range,1,5) = put(min,time5.);
				substr(range,6,1) = " to ";
				substr(range,7,5) = put(max,time5.);
			end;

		/*if n = 0 then
			range = " ";*/

		if frequency = . then
			do;
				frequency = n;
				percent   = round(((n / &obs_cnt) * 100),.01);
			end;

		if _n_ = 1 then
			put "********************** WARNING MESSAGES ***********************";

		if missing(name) then
			do;
				warning = 'POTENTIAL PROBLEM IN DATA:  No occurences for FORMAT category ';
				output _cb_msgs1_;
				put 'Warning: NO OCCURENCES FOR FORMAT CATEGORY: ' FORMAT= DESC=;

				/* If a Format Category does not occur in the data, then delete it, only
				     categories with one or more occurances are listed. */
				delete;
			end;

		if n = 0 and percent=100 then
			do;
				warning = 'POTENTIAL PROBLEM IN DATA:  All values are missing for following variables ';
				put 'Warning: All values are missing for: ' name=;
				output _cb_msgs2_;
			end;

		if min = max and n > 0 and not missing(compress(range)) and not missing(min) and not missing(max) then
			do;
				warning = 'POTENTIAL PROBLEM IN DATA: All Non-Missing values are the same for these variables ';
				put 'Warning: All Non-Missing values are the same for: ' name=;
				output _cb_msgs3_;
			end;

		if missing(label) or upcase(compress(label)) = upcase(compress(name)) then
			do;
				warning = 'POTENTIAL PROBLEM IN DATA: Variable label is missing or matches variable name';
				put 'Warning: Variable label is missing or matches variable name' name= label=;
				output _cb_msgs4_;
			end;

		output _cb_all_data_;

		if lastone = 1 then
			put "******************* END OF WARNING MESSAGES ********************";
		label name      = 'Variable Name'
			label     = 'Label'
			type      = 'Type'
			length    = 'Length'
			type_length = 'Type'
			n         = 'N'                 /* 'Non-Missing Numeric Values' */
		mean      = 'Mean'
			range     = 'Values'
			desc      = 'Frequency Category'
			frequency = 'Frequency'
			percent   = 'Percent'
			cumfrequency = 'Cumulative~Frequency'
			cumpercent   = 'Cumulative~Percent'
			format    = 'Format';
	run;

	data _cb_all_data_(drop=warning) _cb_msgs5_(keep=warning name label format desc range frequency percent cumfrequency cumpercent);
		;
		set _cb_all_data_;
		length WARNING $ 100;

		if compress(range) in ('-', '.-.') and not missing(desc) then
			do;
				WARNING = 'POTENTIAL PROBLEM IN DATA:  Variable has value in data not defined in Format';

				if format not in ('SHOWNUM', '$SHOWALL', '$ANYTEXT', 'ANYMISS') 
					and upcase(compress(desc)) not in ('SASMISSING(.)', 'MISSING(BLANK)') and has_freq_results=1 then
					output _cb_msgs5_;
				range=left(desc);
				desc=' ';
			end;

		if (missing(compress(range, , 'c')) or range = ' ')  and not missing(desc) and upcase(desc) ne 'RANGE' then
			do;
				WARNING = 'POTENTIAL PROBLEM IN DATA:  Variable has value in data not defined in Format';

				if format not in ('SHOWNUM', '$SHOWALL', '$ANYTEXT', 'ANYMISS') 
					and upcase(compress(desc)) not in ('SASMISSING(.)', 'MISSING(BLANK)') and has_freq_results=1 then
					output _cb_msgs5_;
				range=left(desc);
			end;

		output _cb_all_data_;;
	run;

	/** SECTION 10 -  Now print the codebook **/
	%CB_Template(tmp_name=STD);



	%put "CODEBOOK_INFO: Codebook type = " &cb_type;

	proc sort data=_cb_all_data_;
		by name format cnt;
	run;

	%if %sysfunc(exist(work.order)) %then
		%do;
			%put ORDER File Available;

			proc sort data=order;
				by name order;

			data _cb_seq_data_;
				merge order (in=a) _cb_all_data_ (in=b);
				by name;

				if a = 0 then
					put "Order Missing " name=;
			run;

		%end;

	%if %upcase(&var_order)=CUSTOM_ORDER %then
		%do;
			%if %sysfunc(exist(work.custom_order)) %then
				%do;
					%put CODEBOOK ORDER: CUSTOM_ORDER File Available;

					proc sort data=custom_order;
						by name order;

					data _cb_seq_data_;
						merge custom_order (in=a) _cb_all_data_ (in=b);
						by name;

						if a = 0 then
							put "Order Missing " name=;
					run;

				%end;
			%else
				%do;
					%put CODEBOOK ORDER: CUSTOM_ORDER File Not Found, Codebook Ordered Alphabetically;

					data _cb_seq_data_;
						set _cb_all_data_;
						order=1;
				%end;
		%end;
	%else %if %upcase(&var_order) = INTERNAL %then
		%do;
			%put CODEBOOK ORDER: INTERNAL Order of Variables in Data Set Used;

					proc sort data=_cb_var_info_;
						by name;
					run;

					data _cb_seq_data_;
						merge _cb_var_info_(keep=varnum name rename=(varnum=order)) _cb_all_data_(in=b);
						by name;
					run;

		%end;
	%else
		%do;
			%put CODEBOOK ORDER: Codebook Ordered Alphabetically;

			data _cb_seq_data_;
				set _cb_all_data_;
				order=1;
		%end;

	proc sort data=_cb_seq_data_; by order cnt; run;  
	data _cb_seq_data_;
		set _cb_seq_data_;
		order_flag=_n_;

		if compress(desc)="SASmissing(.)" then range=' ';  /*EXCEL turns . into 0 so code to blank for now. */
		if compress(desc)="Missing(blank)" then range=' ';
		if compress(desc)='Range(MintoMax)' and compress(range)='Range(MintoMax)' then range=' ';   
	run;

	/* Create output data set if requested */
	%if "&cb_output" ^= "" %then
		%do;

			data &cb_output (label="Codebook data set for &file1");
				set _cb_seq_data_ (keep= order name label format type_length mean_char order_flag range cnt desc frequency percent CumFrequency CumPercent);
				set_name = "&file1..sas7bdat";
				set_label = "&ds_label  ";
				date_created = "&ds_date";
				num_obs = "&obs_cnt";
				NUM_VAR = "&var_cnt";
				label SET_NAME = "Data set name";
				label SET_LABEL = "Data set label";
				label DATE_CREATED = "Date data set created";
				label NUM_OBS = "Number of observarions in data set";
				label NUM_VAR = "Number of variables in data set";
				label NAME = "Variable name";
				label ORDER = "Variable number order";
				label DESC = "Frequency category text description";
				label FORMAT = "Variable format";
				label LABEL = "Variable label";
				label RANGE = "Values assigned to this category";
				label FREQUENCY = "Frequency";
				label PERCENT = "Percent";
				label CUMFREQUENCY = "Cumulative frequency";
				label CUMPERCENT = "Cumulative percent";
				label CNT = "Order of value defined by RANGE and DESC as it would appear in PROC FREQ output";
				label MEAN_CHAR = "Mean of values for numeric variable";
				label TYPE_LENGTH = "Variable type and length";
				label ORDER_FLAG = "Requested order of variables and values (RANGE/DESC) categories";
			run;

		%end;



	%CB_ODS(open_add_close=OPEN,sname=Codebook);

	ods escapechar='^';
	ods text = " ^S={fontweight=bold fontstyle=italic} Data Set:  &file1..sas7bdat";
	ods text = " ^S={fontweight=bold fontstyle=italic} Label:  &ds_label  ";
	ods text = " ^S={fontweight=bold fontstyle=italic} Date Created:  &ds_date";
	ods text = " ^S={fontweight=bold fontstyle=italic} Number of Observations:  &obs_cnt       Number of Variables:  &var_cnt";

	%if "&organization" ^= "" %then
		%do;
			ods text = "^S={fontweight=bold fontstyle=italic} Organization of Data Set:  &organization";
		%end;

	ods text = " ^{newline 1}";

	%CB_PrintCodebook;

	%if %upcase(&include_warn)=YES %then
		%do;
			/* Value not in FORMAT definition*/
			%CB_IncompFmt(cbwarn=_cb_msgs5_);

			/** OUT OF RANGE  **/
			%CB_PrintOutRange(cbrange=_cb_msgs5_);

			/** ALL HAVE MISSING VALUES **/
			%CB_PrintMissing(cbmiss=_cb_msgs2_);

			/** NO VARIATION IN VALUES **/
			%CB_PrintNoVariation(cbvary=_cb_msgs3_);

			/** Character Variables missing Format Assignment **/
			%CB_PrintMissFmtC(cbnofmtc=_cb_char_var_no_fmt);

			/** Numeric Variables missing Format Assignment **/
			%CB_PrintMissFmtN(cbnofmtn=_cb_num_var_no_fmt);

			/** Variables missing Labels **/
			%CB_PrintNoVarLbl(cbnolbl=_cb_msgs4_);

		%end;
		%CB_ODS( open_add_close=CLOSE,sname=);

/* Delete data sets created by codebook macro */
/*
	proc datasets;
		delete _cb_:;
	run;
*/
%mend;



%Macro CB_PrintCodebook;
	/* PRINT CODEBOOK */
	proc report data=_cb_seq_data_ nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		%if %upcase(&cb_size) = BRIEF %then
			%do;
				column order name label type_length order_flag range cnt desc frequency percent ;
				define order / group noprint;
				define name /group style(column)=[width=15%];
				define label /group flow style(column)=[width=22%];
				define type_length / group center style(column)=[width=5.0%];
				define order_flag/ order noprint;
				define range / group flow center style(column)=[width=10.0%] style(column)={tagattr="format:@"};
				define cnt / group noprint;
				define desc  / group center order flow  style(column)=[width=10.5%] style(column)={tagattr="format:@"};
				define frequency / analysis style(column)=[width=9.8%] format=comma10.0;
				define percent / analysis style(column)=[width=8.6%];
			%end;
		%else
			%do;
				/* default to cb_size=FULL*/
				column order name label format type_length mean_char order_flag range cnt desc frequency percent CumFrequency CumPercent;
				define order / group noprint;
				define name /group style(column)=[width=12%];
				define label /group flow style(column)=[width=22%];
				define type_length / group center style(column)=[width=5.0%];
				define mean_char / group center style(column)=[width=5.8%];
				define order_flag/ order noprint;
				define range / group flow center style(column)=[width=10.0%] style(column)={tagattr="format:@"};
				define format /group style(column)=[width=6.5%];
				define cnt / group noprint;
				define desc  / group center order flow  style(column)=[width=10.5%] style(column)={tagattr="format:@"};
				define frequency / analysis style(column)=[width=5.8%] format=comma10.0;
				define percent / analysis style(column)=[width=4.6%];
				define CumFrequency / analysis style(column)=[width=6.0%] format=comma10.0;
				define CumPercent / analysis style(column)=[width=6.0%];
			%end;

		break after name  /suppress skip;
		break after name  /summarize suppress;
	run;

%mend;

%macro CB_IncompFmt(cbwarn);
	%CB_ODS(open_add_close=ADD,sname=Incomplete Format);
	proc freq data=&cbwarn noprint;
		tables format*desc/list missing nopercent nocum out=_cb_msg5_sumry_;
	run;

	proc report data=_cb_msg5_sumry_ nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM: INCOMPLETE FORMAT~Variable has value not in assigned FORMAT definition.~Check FORMAT definitions.  If correct, check OUT OF RANGE VALUE report."
			format desc count);
		define format /group center style(column)=[width=15%];

		*width=12;
		define desc  / group center order flow  style(column)=[width=40.5%] style(column)={tagattr="format:@"};
		define count /center analysis style(column)=[width=25%] format=comma10.0;
		break after format  /suppress skip;
		break after format  /summarize suppress;
		label desc = "Value not in Format";
		label format = "Format";
		label count = 'Number of Variables with Value ';
	run;

%mend;

%macro CB_PrintOutRange(cbrange);
%CB_ODS(open_add_close=ADD,sname=Out of Range);
	proc sort data=&cbrange;
		by format name desc;
	run;

	proc report data=&cbrange nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM: OUT OF RANGE VALUE~Variable has values not in FORMAT definition. ~Investigate out of range value if FORMAT is correct."
			name label desc format);
		define desc  / group center 'Out of Range Value' order flow  style(column)=[width=10.5%];
		define name /group center style(column)=[width=17%];
		define label /group flow style(column)=[width=30%];
		define format /group center group style(column)=[width=15%];
		break after name  /suppress skip;
		break after name  /summarize suppress;
	run;

%mend;

%macro CB_PrintMissing(cbmiss);
		%CB_ODS(open_add_close=ADD,sname=Missing);	
	proc report data=&cbmiss nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM:  ALL VALUES ARE MISSING FOR THE FOLLOWING VARIABLES"
			name label desc frequency percent);
		define name /group center style(column)=[width=15%];
		define label /group flow style(column)=[width=45%];
		define desc  / group center order flow  style(column)=[width=10.5%];
		define frequency / analysis center style(column)=[width=10%] format=comma10.0;
		define percent / analysis center style(column)=[width=10%];
		break after name  /summarize suppress;
	run;
	;
%mend;

%Macro CB_PrintNoVariation(cbvary);
			%CB_ODS(open_add_close=ADD,sname=Unvarying);
proc report data=&cbvary nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ( "POTENTIAL PROBLEM: NUMERIC VARIABLES WITH NO VARIATION IN RESPONSE~All Non-missing values are the same for these variables"
			name label desc min max frequency percent);
		define name /group center style(column)=[width=12%];
		define label /group flow style(column)=[width=40%];
		define desc  / group center order flow  style(column)=[width=10.5%] style(column)={tagattr="format:@"};
		define min/analysis center style(column) = [width=6%];
		define max/analysis center style(column) = [width=6%];
		define frequency / analysis center style(column)=[width=6%] format=comma10.0;
		define percent / analysis center style(column)=[width=6%];
		label min="Minimum";
		label max="Maximum";
		break after name  /suppress skip;
		break after name  /summarize suppress;
	run;

%mend;

%macro CB_PrintMissFmtC(cbnofmtc);
	%CB_ODS(open_add_close=ADD,sname=Missing Format (Character));
	proc report data=&cbnofmtc nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM: CHARACTER VARIABLES NOT ASSIGNED USER FORMAT~Codebook combines all values in the category 'BLANK, TEXT, OR VALUE SUPPLIED'~Try format with categories 'BLANK', 'TEXT OR VALUE SUPPLIED'"
			name type length label);
		define name /group center style(column)=[width=25%];
		define type /group center style(column) = [width=10%];
		define length/group center style(column) = [width=10%];
		define label /group center flow style(column)=[width=40%];
	run;

%mend;

%macro CB_PrintMissFmtN(cbnofmtn);
			%CB_ODS(open_add_close=ADD,sname=Missing Format (Numeric));
	proc report data=&cbnofmtn nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM: NUMERIC VARIABLES NOT ASSIGNED USER FORMAT~Codebook uses categories 'Range', 'SAS Missing (.)'~Try format with categories 'Valid Range', 'Missing'~Values outside of Valid Range will be identified in Out of Range Report."
			name type length label);
		define name /group center style(column)=[width=25%];
		define type /group center style(column) = [width=10%];
		define length/group center style(column) = [width=10%];
		define label /group center flow style(column)=[width=40%];
	run;

%mend;

%macro CB_PrintNoVarLbl(cbnolbl);
	%CB_ODS(open_add_close=ADD,sname=Undefined Var Label);
	proc sort data=&cbnolbl(keep=warning name label) nodupkey;
		by warning name label;
	run;

	data &cbnolbl;
	set &cbnolbl;
	length potential_problem $35;
	if missing(label) then potential_problem='Missing variable label';
	else potential_problem='Label identical to variable name';
	label potential_problem='Potential Problem';
	run;

	proc report data=&cbnolbl nocenter nowindows headline wrap split='~'  missing
		style(header)=[color=black backgroundcolor=very light grey ]
		style(summary)=[color=very light grey backgroundcolor=very light grey fontfamily="Times Roman" fontsize=1pt textalign=r];;
		column ("POTENTIAL PROBLEM: UNDEFINED VARIABLE LABEL ~Variable is missing label or has label matching variable name."
			name label  potential_problem);
		define name /group center style(column)=[width=25%];
		define label/group center style(column)=[width=25%];
		define potential_problem /group center flow style(column)=[width=35%];
	run;

%mend;

/*--------------------------------------------------------------------------------------------------------------------------;
 * Program	max_lengths.sas                                                                                            ;
 * Purpose	Creates a SAS macro variable &char_lengths that has the maximum length of two characters variables that    ;
 *		have the same name but different lengths in two data sets.  Use this variable in a length statement:       ;
 *			length &char_lengths ;                                                                             ;
 * 		in a data step so that data is not truncated.                                                              ;
 * date		June, 22, 2017                                                                                             ;
 *                                                                                                                          ;
 * Usage:  	%max_lengths(set1, set2);                                                                                  ;
 * Parameters:	set1 = name of first data set with character variables                                                     ;
 *		set2 = name of second data set with character variables                                                    ;
 *--------------------------------------------------------------------------------------------------------------------------*/
%macro max_lengths(set1, set2);

	proc contents data=&set1 noprint out=_cb_out1_(keep=name type length where=(type=2));

	proc contents data=&set2 noprint out=_cb_out2_(keep=name type length where=(type=2));

	data _cb_out1_;
		set _cb_out1_(drop=type);
		name=upcase(name);
		rename length=length1;
	run;

	data _cb_out2_;
		set _cb_out2_(drop=type);
		name=upcase(name);
		rename length=length2;
	run;

	proc sort data=_cb_out1_;
		by name;
	run;

	proc sort data=_cb_out2_;
		by name;
	run;

	data _cb_lengths_;
		merge _cb_out1_ _cb_out2_;
		by name;
		length_info=compress(Name) || ' $ ' || max(length1, length2);
	run;

	%global char_lengths;
	%let char_lengths=;

	Proc SQL;
		Select length_info into :char_lengths separated by ' ' from work._cb_lengths_;
	quit;

%mend max_lengths;

/****************************************************************************************************************************************
 * program       find_dups.sas                                                                                                          ;
 * function      Find observations with duplicate ID values.                                                                            ;
 * date          July 12, 2017                                                                                                          ;
 * usage                                                                                                                                ;
 *              %find_dups(dataset, one_rec_per);                                                                                          ;
 *               where:                                                                                                                 ;
 *                    dataset= Name of Data set being examine for duplicate values of Identification Variables                          ;
 *                    one_rec_per = Variable combination that specifies the duplicate Identification variables.  For example,           ;
 *                                  if the data set should have only one record for every person (CASEID) for every wave of data (WAVE) ;
 *                                  then one_rec_per = CASEID*WAVE.                                                                     ;
 ***************************************************************************************************************************************/
%macro find_dups(dataset, one_rec_per);

	proc freq data=&dataset noprint;
		tables &one_rec_per/list missing out=_dup_check_;
		format _all_;
	run;

	proc freq data=_dup_check_;
		title4 "Find records with duplicate values of Identification Variables in Data Set";
		title5 "Data set being examined:  &dataset";
		title6 "Identification Variables:  &one_rec_per";
		title7 "There should be only one record for every unique value of &one_rec_per";
		tables count/list missing;
		label count="COUNT = Count of records with identical values of &one_rec_per";
		format _all_;
	run;

	proc freq data=_dup_check_;
		title4 "Values of Identification Variables with duplicate values of &one_rec_per in data set &dataset";
		title5 "COUNT = Count of records with identical values of &one_rec_per";
		where count>1;
		tables &one_rec_per*count /list missing;
		format _all_;
	run;

	title3 ' ';
%mend find_dups;
%macro CB_ODS(open_add_close, sname);
%if %upcase(&cb_type) = XML %then
	%do;
		%if %upcase(&open_add_close) = OPEN %then
			%do;
				ods tagsets.excelxp file="&cb_file" style=custom options(embedded_titles="yes" gridlines='on' sheet_name="&Sname"  orientation='landscape');
			%end;

		%if %upcase(&open_add_close) = ADD %then
			%do;
				ods tagsets.excelxp options(embedded_titles="yes" gridlines='on' sheet_name="&Sname" orientation='landscape' );
			%end;

		%if %upcase(&open_add_close) = CLOSE %then
			%do;
				ods tagsets.excelxp close;
				ods listing;
			%end;
	%end;

%if %upcase(&cb_type) = XLS or  %upcase(&cb_type) = XLSX  %then
	%do;
		%if %upcase(&open_add_close) = OPEN %then
			%do;
				ods EXCEL file="&cb_file" style=custom options(embedded_titles="yes" gridlines='on' sheet_name="&Sname"  orientation='landscape');
			%end;

		%if %upcase(&open_add_close) = ADD %then
			%do;
				ods EXCEL  style=custom options(embedded_titles="yes" gridlines='on' sheet_name="&Sname" orientation='landscape' );
			%end;

		%if %upcase(&open_add_close) = CLOSE %then
			%do;
				ods excel close;
				ods listing;
			%end;
	%end;

%if %upcase(&cb_type) = PDF %then
	%do;
		%if %upcase(&open_add_close) = OPEN %then
			%do;
				options orientation=landscape;
				ods pdf file="&cb_file"  style=custom;
			%end;

		%if %upcase(&open_add_close) = CLOSE %then
			%do;
				ods pdf close;
				ods listing;
			%end;
	%end;

%if %upcase(&cb_type) = RTF %then
	%do;
		%if %upcase(&open_add_close) = OPEN %then
			%do;
				options orientation=landscape;
				ods rtf file="&cb_file"  style=custom;
			%end;

		%if %upcase(&open_add_close) = CLOSE %then
			%do;
				ods rtf close;
				ods listing;
			%end;
	%end;
	title3 ' ';

%mend;
%macro CB_template(tmp_name);
	proc template;
		define style Styles.custom;
			parent=Styles.RTF;
			replace fonts /
				'TitleFont' = ("Times",9pt,Bold Italic) 
				'TitleFont2' = ("Times",9pt,Bold Italic) 
				'StrongFont' = ("Times",9pt,Bold)       
				'EmphasisFont' = ("Times",9pt,Italic)    
				'headingEmphasisFont' = ("Times",9pt,Bold Italic)
				'headingFont' = ("Times",9pt,Bold)
				'docFont' = ("Times",9pt) 
				'footFont' = ("Times",9pt) 
				'FixedEmphasisFont' = ("Times",9pt,Italic)
				'FixedStrongFont' = ("Times",9pt,Bold)
				'FixedHeadingFont' = ("Times",9pt,Bold)
				'BatchFixedFont' = ("Times",9pt)
				'FixedFont' = ("Times",9pt);
			replace Body from Document /
				bottommargin = 0.25in
				topmargin = 0.25in
				rightmargin = 0.25in
				leftmargin = 0.25in;
			replace color_list /
				'link' = blue  /*links */

			'bgH' = white  /*row and column header background*/
			'fg' = black   /*text color*/
			'bg' = white;  /* page background color */
			replace Table from Output/
			frame=hsides
			rules=groups /*all*/
			cellpadding=3pt
			cellspacing=0pt
			borderwidth=.75pt;
		end;
	run;
%mend;

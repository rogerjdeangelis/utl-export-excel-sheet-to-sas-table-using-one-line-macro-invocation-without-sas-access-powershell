# utl-export-excel-sheet-to-sas-table-using-one-line-macro-invocation-without-sas-access-powershell
Export excel sheet to sas table using one line macro invocation without sas access powershell
    %let pgm=utl-export-excel-sheet-to-sas-table-using-one-line-macro-invocation-without-sas-access-powershell;

    %stop_submission;

    Export excel sheet to sas table using one line macro invocation without sas access powershell

      Contents
         1 usage w powershell macro wrapper
         2 drop down powershell

    github
    https://tinyurl.com/2d92a73n
    https://github.com/rogerjdeangelis/utl-export-excel-sheet-to-sas-table-using-one-line-macro-invocation-without-sas-access-powershell

    SOAPBOX ON

    ISSUE
      SAS cannot directly create a sas table from a excel sheet
    SOAPBOX OFF

    /**************************************************************************************************************************/
    /* INPUT                                  |  PROCESS                                     | OUTPUT                         */
    /* =====                                  |  =======                                     | ======                         */
    /* d:/xls/class.xlsx                      |  1 USAGE W POWERSHELL MACRO WRAPPER          | NAME    SEX AGE HEIGHT WEIGHT  */
    /*                                        |  ==================================          |                                */
    /* ----------------------+                |                                              | Alfred  M   14  69     112.5   */
    /* | A1| fx    |NAME     |                |  %utl_getsheet(                              | Alice   F   13  56.5   84      */
    /* --------------------------------------+|      wb=d:\xls\class.xlsx                    | Barbara F   13  65.3   98      */
    /* [_] |    A    | B | C |    D  |    E  ||     ,sheet=class                             | Carol    F   14  62.8   102.5  */
    /* --------------------------------------||     ,csv=d:\csv\class.csv                    | Henry   M   14  63.5   102.5   */
    /*  1  | NAME    |SEX|AGE| HEIGHT| WEIGHT||     ,table=classout                          |                                */
    /*  -- |---------+---+---+-------+-------||     );                                       |                                */
    /*  2  |  Alfred | M | 14| 69    | 112.5 ||                                              |                                */
    /*  -- |---------+---+---+-------+-------||  MACRO                                       |                                */
    /*  3  |  Alice  | F | 13| 56.5  | 84    ||                                              |                                */
    /*  -- |---------+---+---+-------+-------||  %macro utl_getsheet(                        |                                */
    /*  4  |  Barbara| F | 13| 65.3  | 98    ||    wb=d:\xls\class.xlsx                      |                                */
    /*  -- |---------+---+---+-------+-------||   ,sheet=class                               |                                */
    /*  5  |  Carol  | F | 14| 62.8  | 102.5 ||   ,csv=d:\csv\class.csv                      |                                */
    /*  -- |---------+---+---+-------+-------||   ,table=classout                            |                                */
    /*  6  |  Henry  | M | 14| 63.5  | 102.5 ||   ) /des="Export excel sheet to sas table";  |                                */
    /*  -- |---------+---+---+-------+-------||                                              |                                */
    /* [CLASS}                                |   %utlfkil(d:/xls/class.xlsx);               |                                */
    /*                                        |                                              |                                */
    /* data have;informat                     |   %utl_submit_ps64x(resolve(                 |                                */
    /* NAME $8.                               |      'import-Excel -Path "&wb"               |                                */
    /* SEX $1.                                |        -WorksheetName "&sheet" | Export-Csv  |                                */
    /* AGE 8.                                 |        -Path "&csv" -NoTypeInformation'));   |                                */
    /* HEIGHT 8.                              |   dm "";                                     |                                */
    /* WEIGHT 8.                              |   dm "dimport '&csv' &table replace";        |                                */
    /* ;input                                 |                                              |                                */
    /* NAME SEX AGE                           |  %mend utl_getsheet;                         |                                */
    /*  HEIGHT WEIGHT;                        |                                              |                                */
    /* cards4;                                |  2 DROP DOWN TO POWERSHELL                   |                                */
    /* Alfred M 14 69 112.5                   |  =========================                   |                                */
    /* Alice F 13 56.5 84                     |                                              |                                */
    /* Barbara F 13 65.3 98                   |  see below and github                        |                                */
    /* Carol F 14 62.8 102.5                  |                                              |                                */
    /* Henry M 14 63.5 102.5                  |                                              |                                */
    /* ;;;;                                   |                                              |                                */
    /* run;quit;                              |                                              |                                */
    /**************************************************************************************************************************/

    /*___        _                       _                                                       _          _ _
    |___ \    __| |_ __ ___  _ __     __| | _____      ___ __   _ __   _____      _____ _ __ ___| |__   ___| | |
      __) |  / _` | `__/ _ \| `_ \   / _` |/ _ \ \ /\ / / `_ \ | `_ \ / _ \ \ /\ / / _ \ `__/ __| `_ \ / _ \ | |
     / __/  | (_| | | | (_) | |_) | | (_| | (_) \ V  V /| | | || |_) | (_) \ V  V /  __/ |  \__ \ | | |  __/ | |
    |_____|  \__,_|_|  \___/| .__/   \__,_|\___/ \_/\_/ |_| |_|| .__/ \___/ \_/\_/ \___|_|  |___/_| |_|\___|_|_|
                            |_|                                |_|
    */

    %macro utl_submit_ps64x(
          pgm
         ,return=  /* name for the macro variable from Powershell */
         )/des="Semi colon separated set of Powershell commands - drop down to Powershell";
      /*
          %let pgm='Get-Content -Path d:/txt/back.txt | Measure-Object -Line | clip;';
      */
      * write the program to a temporary file;
      filename py_pgm "%sysfunc(pathname(work))/py_pgm.ps1" lrecl=32766 recfm=v;
      data _null_;
        length pgm  $32755 cmd $1024;
        file py_pgm ;
        pgm=&pgm;
        semi=countc(pgm,';');
          do idx=1 to semi;
            cmd=cats(scan(pgm,idx,';'));
            if cmd=:'. ' then
               cmd=trim(substr(cmd,2));
             put cmd $char384.;
             putlog cmd $char384.;
          end;
      run;quit;
      %let _loc=%sysfunc(pathname(py_pgm));
      %put &_loc;
      filename rut pipe  "powershell.exe -executionpolicy bypass -file &_loc ";
      data _null_;
        file print;
        infile rut;
        input;
        put _infile_;
        putlog _infile_;
      run;
      filename rut clear;
      filename py_pgm clear;
      * use the clipboard to create macro variable;
      %if "&return" ^= "" %then %do;
        filename clp clipbrd ;
        data _null_;
         length txt $200;
         infile clp;
         input;
         putlog "*******  " _infile_;
         call symputx("&return",_infile_,"G");
        run;quit;
      %end;
    %mend utl_submit_ps64x;

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */

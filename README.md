# utl-update-an-existing-excel-named-range-R-python-sas
Update an existing excel named range, "males" R Python SAS
    Update an existing excel named range, "males" R Python SAS

    github
    https://tinyurl.com/sg5ohbp
    https://github.com/rogerjdeangelis/utl-update-an-existing-excel-named-range-R-python-sas

    I don't think this is easily done in vanilla SAS.
    Python and R can do it.

    SQL update, datastep modify and proc append will fail with this

    This table already exists, or there is a name conflict with an existing object.
    This table will not be replaced.
    This engine does not support the REPLACE option.

    Maybe the replace option in proc export will work.

      We will update existing named range "males" with new data

         Methods

          a. R

            1.  Have SAS create an excel workbook with two sheets("males","females")
                and two named ranges,("males","females")
            2.  Have R remove sheet "males" and named range "males"
            3.  Have R create Sheet "males" and named range "males"
            4.  Have R update named range "males" with new data

          b. SAS with a little help from R

            1.  Have R remove the sheet "males" and then named range "males"
            2.  Have SAS create named range "males", sheet name "males"
                and populate the named range "males"

          c. Python (cleanest solution. Do not have to remove sheets and named ranges)

             1. Just read loop though the updates and load into the named range

    SAS Forum
    https://tinyurl.com/u9599xw
    https://communities.sas.com/t5/New-SAS-User/SAS-Studio-Writing-to-a-named-range-in-Excel/m-p/631538

    Related Repos
    https://github.com/rogerjdeangelis?tab=repositories&q=excel+in%3Aname&type=&language=

    *          ____
      __ _    |  _ \
     / _` |   | |_) |
    | (_| |_  |  _ <
     \__,_(_) |_| \_\
     _                    _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;


    %utlfkil(d:/xls/xis.xlsx);

    libname xel "d:/xls/xis.xlsx";

    data xel.females xel.males;
      set sashelp.class ;
      if sex="M" then output xel.males;
      else output xel.females;
    run;quit;

    libname xel clear;

    * UPDATES FOR MALES NAMED RANGE;

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.males;
      set sashelp.class(where=(sex="M"));
      sex=substr(name,1,1);
      age=age+1000;
      name=left(reverse(name));
    run;quit;


    SHEET CLASS IN WORKBOOK D:/XLS/XIS.XLSX

    NAMED RANGE MALES

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
    1 |  NAME      |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
    2 | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
    N | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

    [MALES]


    SAS UPDATES FOR EXCEL NAMED RANGE "MALES"


    SD1.MALES total obs=10

      NAME       SEX     AGE    HEIGHT    WEIGHT

      semaJ       J     1012     57.3       83.0
      samohT      T     1011     57.5       85.0
      nhoJ        J     1012     59.0       99.5
      yerffeJ     J     1013     62.5       84.0
      yrneH       H     1014     63.5      102.5
      treboR      R     1012     64.8      128.0
      mailliW     W     1015     66.5      112.0
      dlanoR      R     1015     67.0      133.0
      derflA      A     1014     69.0      112.5
      pilihP      P     1016     72.0      150.0
     ...
    *            _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| '_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    ;

    NAMED RANGE MALES

       NOTE NAME AND SEX ARE DIFFERENT

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
    1 |  NAME      |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
    2 | DERFLA     |    A       |  1014      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
    N | MILLIWM    |    W       |  1015      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

    [MALES]

    %utl_submit_r64('
    library(haven);
    library(XLConnect);
    males<-read_sas("d:/sd1/males.sas7bdat");
    wb <- loadWorkbook("d:/xls/xis.xlsx");
    removeSheet(wb,"males");
    removeName(wb,"males");
    createSheet(wb,"males");
    createName(wb, name = "males", formula = "males!$A$1",overwrite=TRUE);
    writeNamedRegion(wb, males, name = "males");
    saveWorkbook(wb);
    ');

    *_                         ___     ____
    | |__     ___  __ _ ___   ( _ )   |  _ \
    | '_ \   / __|/ _` / __|  / _ \/\ | |_) |
    | |_) |  \__ \ (_| \__ \ | (_>  < |  _ <
    |_.__(_) |___/\__,_|___/  \___/\/ |_| \_\
     _                    _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    %utlfkil(d:/xls/xis.xlsx);

    libname xel "d:/xls/xis.xlsx";

    data xel.females xel.males;
      set sashelp.class ;
      if sex="M" then output xel.males;
      else output xel.females;
    run;quit;

    libname xel clear;

    * UPDATES FOR MALES NAMED RANGE;

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.males;
      set sashelp.class(where=(sex="M"));
      sex=substr(name,1,1);
      age=age+1000;
      name=left(reverse(name));
    run;quit;

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    %utl_submit_r64('
    library(haven);
    library(XLConnect);
    males<-read_sas("d:/sd1/males.sas7bdat");
    wb <- loadWorkbook("d:/xls/xis.xlsx");
    removeSheet(wb,"males");
    removeName(wb,"males");
    saveWorkbook(wb);
    ');

    libname xel "d:/xls/xis.xlsx";

    data xel.males;
      set sashelp.class(where=(sex="M"));
      sex=substr(name,1,1);
      age=age+1000;
      name=left(reverse(name));
    run;quit;

    libname xel clear;

    *         ____        _   _
      ___    |  _ \ _   _| |_| |__   ___  _ __
     / __|   | |_) | | | | __| '_ \ / _ \| '_ \
    | (__ _  |  __/| |_| | |_| | | | (_) | | | |
     \___(_) |_|    \__, |\__|_| |_|\___/|_| |_|
     _                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    %utlfkil(d:/xls/xis.xlsx);

    libname xel "d:/xls/xis.xlsx";

    data xel.females xel.males;
      set sashelp.class ;
      if sex="M" then output xel.males;
      else output xel.females;
    run;quit;

    libname xel clear;

    * UPDATES FOR MALES NAMED RANGE;

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.males;
      set sashelp.class(where=(sex="M"));
      sex=substr(name,1,1);
      age=age+1000;
      name=left(reverse(name));
    run;quit;


    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|
    ;

    * this works;
    %utl_submit_py64("
    from openpyxl.utils.dataframe import dataframe_to_rows;
    from openpyxl import Workbook;
    from openpyxl import load_workbook;
    from sas7bdat import SAS7BDAT;
    with SAS7BDAT('d:/sd1/males.sas7bdat') as m:;
    .   clas = m.to_data_frame();
    wb = load_workbook(filename='d:/xls/xis.xlsx', read_only=False);
    ws = wb.get_sheet_by_name('males');
    rows = dataframe_to_rows(clas);
    for r_idx in range(len(clas)):;
    .   for c_idx in range(len(clas.columns)):;
    .        c=c_idx+1;
    .        r=r_idx+1;
    .        ws.cell(row=r_idx+2, column=c_idx+1,value=clas.iloc[r-1,c-1]);
    wb.save('d:/xls/xis.xlsx');
    ");




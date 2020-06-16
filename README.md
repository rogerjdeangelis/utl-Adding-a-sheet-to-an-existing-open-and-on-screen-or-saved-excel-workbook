# utl-adding-a-sheet-to-an-existing-open-and-on-screen-or-saved-excel-workbook
Adding a sheet to an existing open and on screen or saved excel workbook                                                                
    Adding a sheet to an existing open and on screen or saved excel workbook                                  
                                                                                                              
    Another reason to avoid proc export to excel                                                              
                                                                                                              
    I use the more functional Excel 2010 Plus not Office 365.                                                 
    May or may not work in Office 365?                                                                        
                                                                                                              
                                                                                                              
        Method (both allow you to add worksheets to an existing workbook)                                     
                                                                                                              
          a.  Open excel workbook and click on the review tab then select share workbook                      
              Keep workbook open and on screen                                                                
                                                                                                              
          b.  Open excel workbook and click on the review tab then select share workbook                      
              Save work book and close excel                                                                  
              Proc export may work - not sure - and don't care                                                
    *                _     _                                                                                  
     _ __  _ __ ___ | |__ | | ___ _ __ ___                                                                    
    | '_ \| '__/ _ \| '_ \| |/ _ \ '_ ` _ \                                                                   
    | |_) | | | (_) | |_) | |  __/ | | | | |                                                                  
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|                                                                  
    |_|                                                                                                       
    ;                                                                                                         
    Here is the problem                                                                                       
                                                                                                              
    Existing workbook d:/xls/problem.xlsx is open on screen                                                   
    I want to add sheet heart to the workbook.                                                                
    There may be some clever option to make this wor?                                                         
                                                                                                              
    proc export                                                                                               
       data=sashelp.heart                                                                                     
       dbms=xlsx                                                                                              
       outfile="d:/xls/problem.xlsx";                                                                         
    run;                                                                                                      
                                                                                                              
    ERROR: Sheet (problem) in workbook (d:\xls\problem.xlsx)                                                  
          already exists. Specify REPLACE option to overwrite it.                                             
    ERROR: Export unsuccessful.  See SAS Log for details.                                                     
                                                                                                              
    LESS IS MORE (This does not mean to remove usefull functionality just the                                 
    massive bloatware clutter with problematic functionality or not                                           
    needed funtionality)                                                                                      
                                                                                                              
    sas28rlxxlr82sas                                                                                          
    *_                   _                                                                                    
    (_)_ __  _ __  _   _| |_                                                                                  
    | | '_ \| '_ \| | | | __|                                                                                 
    | | | | | |_) | |_| | |_                                                                                  
    |_|_| |_| .__/ \__,_|\__|                                                                                 
            |_|                                                                                               
    ;                                                                                                         
                                                                                                              
    %utlfkil(d:/xls/class.xlsx);                                                                              
                                                                                                              
    libname xel "d:/xls/class.xlsx";                                                                          
                                                                                                              
    data xel.class;                                                                                           
       set sashelp.class;                                                                                     
    run;quit;                                                                                                 
                                                                                                              
    libname xel clear;                                                                                        
                                                                                                              
    * open the the excel workbook and turn on sharing                                                         
      click on the review tab then select share workbook;                                                     
                                                                                                              
                                                                                                              
    "d:/xls/class.xlsx"                                                                                       
                                                                                                              
          +--------------------------------------------------------------+                                    
          |     A      |    B       |     C      |    D       |    E     |                                    
          +--------------------------------------------------------------+                                    
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT  |                                    
          +------------+------------+------------+------------+----------+                                    
       2  | Alice      |    f      |    12       |   88       |   56     |                                    
          +------------+------------+------------+------------+----------+                                    
       3  | Mary       |    Mf      |    16      |    11      |   67     |                                    
          +------------+------------+------------+------------+----------+                                    
       4  | Tom        |    M       |    15      |    99      |   33     |                                    
          +------------+------------+------------+------------+----------+                                    
       5  | John       |    M       |    14      |    77      |   90     |                                    
          +------------+------------+------------+------------+----------+                                    
       6  | Jane       |    F       |    13      |    44      |   67     |                                    
          +------------+------------+------------+------------+----------+                                    
                                                                                                              
         [CLASS]                                                                                              
                                                                                                              
                                                                                                              
    *            _               _                                                                            
      ___  _   _| |_ _ __  _   _| |_                                                                          
     / _ \| | | | __| '_ \| | | | __|                                                                         
    | (_) | |_| | |_| |_) | |_| | |_                                                                          
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                         
                    |_|                                                                                       
    ;                                                                                                         
                                                                                                              
    * Note the additional MALES Sheet;                                                                        
                                                                                                              
    "d:/xls/class.xlsx"                                                                                       
                                                                                                              
          +--------------------------------------------------------------+                                    
          |     A      |    B       |     C      |    D       |    E     |                                    
          +--------------------------------------------------------------+                                    
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT  |                                    
          +------------+------------+------------+------------+----------+                                    
       2  | Alice      |    f      |    12       |   88       |   56     |                                    
          +------------+------------+------------+------------+----------+                                    
       3  | Mary       |    Mf      |    16      |    11      |   67     |                                    
          +------------+------------+------------+------------+----------+                                    
       4  | Tom        |    M       |    15      |    99      |   33     |                                    
          +------------+------------+------------+------------+----------+                                    
       5  | John       |    M       |    14      |    77      |   90     |                                    
          +------------+------------+------------+------------+----------+                                    
       6  | Jane       |    F       |    13      |    44      |   67     |                                    
          +------------+------------+------------+------------+----------+                                    
                                                                                                              
         [CLASS] [MALES] [FEMALES]                                                                            
                                                                                                              
    *          _       _   _                                                                                  
     ___  ___ | |_   _| |_(_) ___  _ __                                                                       
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                                      
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                     
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                     
                                                                                                              
    ;                                                                                                         
                                                                                                              
    * AFTER YOU RUN THE CODE BELOW DO NOT SAVE THE OPEN WORKBOOK                                              
      JUST EXIT OUT. EXCEL WILL NOT AS YOU TO SAVE.                                                           
                                                                                                              
    libname xel "d:/xls/class.xlsx";                                                                          
                                                                                                              
    data xel.males;                                                                                           
       set sashelp.class(where=(sex="M"));                                                                    
    run;quit;                                                                                                 
                                                                                                              
    libname xel clear;                                                                                        
                                                                                                              
                                                                                                              
    libname xel "d:/xls/class.xlsx";                                                                          
                                                                                                              
    data xel.females;                                                                                         
       set sashelp.class(where=(sex="F"));                                                                    
    run;quit;                                                                                                 
                                                                                                              
    libname xel clear;                                                                                        
                                                                                                              
                                                                                                              
                                                                                                              

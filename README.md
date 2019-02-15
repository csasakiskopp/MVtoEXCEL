# MVtoEXCEL
UniData UniBasic language tool for building openXML spreadsheets (.xlsx)

Requirements:

UniData v 7.2 or higher on Linux

To Install on a system running UniData:

Place code in a directory on your Linux system running UniData DBMS. Set a VOC pointer to the install file named MV.TO.EXCEL. The Templates directory should be given its own voc pointer of MVTX.TEMPLATE.  

General usage:

Put an include at the top of your BASIC program as below-
    $INCLUDE MV.TO.EXCEL PUBLIC.INCLUDE

This will give the program access to the spreadsheet building functions:

    MAKE.WB(OWNER)                         ;*Activates a workbook. Will clear any existing workbook
    WB.ACTIVE                              ;*Returns TRUE/FALSE as to whether a Workbook is active
    ADD.SHEET(SHEETNAME,OPTS)              ;*Add a sheet to a workbook. Returns a sheet index number
    ADD.ROW(SHEET.INDEX,CELL.DATA)         ;*Add a row of spreadsheet data to a worksheet
    ADD.STYLE(STYLE.DATA)                  ;*Add a defined cell style. Returns a style index number
    ADD.TABLE(SHEET.INDEX,TABLE.DATA)      ;*Add a defined table area to a worksheet
    WRITE.WB(BOOK.NAME,DIR.LOCATION)       ;*Write a completed workbook to a location in linux
    ADD.TABLE(SHEET.IDX,TABLE.DATA)        ;*Add a defined data-area to a worksheet
    GET.ROW.CNT(SHEET.IDX)  

Design Philosophy:

Most UniBasic applications are written in a highly procedural style. This one instead uses Functions to compartmentalize the discrete tasks necessary to assemble an openXML document. Inside each function, further compartmentalization occurs through the use of internal subroutines, dividing tasks up simple steps. As a whole, the function set acts as though they were a set of methods acting on an object. The object takes the form of a shared common space with pre-defined Equate values. The danger here is that UniBasic does not support objects, so the object-like nature of this commons space, and the rules that make it function are encapsulated in the MVTX.EQUATES file UniBasic does not have scoping beyond the individual program/function, so MVtoEXCEL relies on the programmer to adhere to and preserve these structures in their application. 

The primary engine behind these functions is uniBasic's XDOM extension. Because this extension does not handle XML with any particular elegance, starter templates are used to facilitate namespacing. Each template is loaded into the common area at the start of workbook formation and added to with each subsequent function. 

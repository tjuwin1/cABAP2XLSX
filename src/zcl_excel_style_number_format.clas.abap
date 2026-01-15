class zcl_excel_style_number_format definition
  public
  final
  create public .

  public section.

    types:
      begin of t_num_format,
        id     type string,
        format type ref to zcl_excel_style_number_format,
      end of t_num_format .
    types:
      t_num_formats type hashed table of t_num_format with unique key id .

*"* public components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
    constants c_format_numc_std type zif_excel_data_decl=>zexcel_number_format value 'STD_NDEC'. "#EC NOTEXT
    constants c_format_date_std type zif_excel_data_decl=>zexcel_number_format value 'STD_DATE'. "#EC NOTEXT
    constants c_format_currency_eur_simple type zif_excel_data_decl=>zexcel_number_format value '[$EUR ]#,##0.00_-'. "#EC NOTEXT
    constants c_format_currency_usd type zif_excel_data_decl=>zexcel_number_format value '$#,##0_-'. "#EC NOTEXT
    constants c_format_currency_usd_simple type zif_excel_data_decl=>zexcel_number_format value '"$"#,##0.00_-'. "#EC NOTEXT
    constants c_format_currency_simple type zif_excel_data_decl=>zexcel_number_format value '$#,##0_);($#,##0)'. "#EC NOTEXT
    constants c_format_currency_simple_red type zif_excel_data_decl=>zexcel_number_format value '$#,##0_);[Red]($#,##0)'. "#EC NOTEXT
    constants c_format_currency_simple2 type zif_excel_data_decl=>zexcel_number_format value '$#,##0.00_);($#,##0.00)'. "#EC NOTEXT
    constants c_format_currency_simple_red2 type zif_excel_data_decl=>zexcel_number_format value '$#,##0.00_);[Red]($#,##0.00)'. "#EC NOTEXT
    constants c_format_date_datetime type zif_excel_data_decl=>zexcel_number_format value 'd/m/y h:mm'. "#EC NOTEXT
    "! Deprecated. Do not use this one, its value is dd/mm/yy, instead use the constant *_ddmmyyyy_new
    constants c_format_date_ddmmyyyy type zif_excel_data_decl=>zexcel_number_format value 'dd/mm/yy'. "#EC NOTEXT
    constants c_format_date_ddmmyyyy_new type zif_excel_data_decl=>zexcel_number_format value 'dd/mm/yyyy'. "#EC NOTEXT
    constants c_format_date_ddmmyy type zif_excel_data_decl=>zexcel_number_format value 'dd/mm/yy'. "#EC NOTEXT
    constants c_format_date_ddmmyyyydot type zif_excel_data_decl=>zexcel_number_format value 'dd\.mm\.yyyy'. "#EC NOTEXT
    constants c_format_date_dmminus type zif_excel_data_decl=>zexcel_number_format value 'd-m'. "#EC NOTEXT
    constants c_format_date_dmyminus type zif_excel_data_decl=>zexcel_number_format value 'd-m-y'. "#EC NOTEXT
    constants c_format_date_dmyslash type zif_excel_data_decl=>zexcel_number_format value 'd/m/y'. "#EC NOTEXT
    constants c_format_date_myminus type zif_excel_data_decl=>zexcel_number_format value 'm-y'. "#EC NOTEXT
    constants c_format_date_time1 type zif_excel_data_decl=>zexcel_number_format value 'h:mm AM/PM'. "#EC NOTEXT
    constants c_format_date_time2 type zif_excel_data_decl=>zexcel_number_format value 'h:mm:ss AM/PM'. "#EC NOTEXT
    constants c_format_date_time3 type zif_excel_data_decl=>zexcel_number_format value 'h:mm'. "#EC NOTEXT
    constants c_format_date_time4 type zif_excel_data_decl=>zexcel_number_format value 'h:mm:ss'. "#EC NOTEXT
    constants c_format_date_time5 type zif_excel_data_decl=>zexcel_number_format value 'mm:ss'. "#EC NOTEXT
    constants c_format_date_time6 type zif_excel_data_decl=>zexcel_number_format value 'h:mm:ss'. "#EC NOTEXT
    constants c_format_date_time7 type zif_excel_data_decl=>zexcel_number_format value 'i:s.S'. "#EC NOTEXT
    constants c_format_date_time8 type zif_excel_data_decl=>zexcel_number_format value 'h:mm:ss@'. "#EC NOTEXT
    constants c_format_date_xlsx14 type zif_excel_data_decl=>zexcel_number_format value 'mm-dd-yy'. "#EC NOTEXT
    constants c_format_date_xlsx15 type zif_excel_data_decl=>zexcel_number_format value 'd-mmm-yy'. "#EC NOTEXT
    constants c_format_date_xlsx16 type zif_excel_data_decl=>zexcel_number_format value 'd-mmm'. "#EC NOTEXT
    constants c_format_date_xlsx17 type zif_excel_data_decl=>zexcel_number_format value 'mmm-yy'. "#EC NOTEXT
    constants c_format_date_xlsx22 type zif_excel_data_decl=>zexcel_number_format value 'm/d/yy h:mm'. "#EC NOTEXT
    constants c_format_date_yymmdd type zif_excel_data_decl=>zexcel_number_format value 'yymmdd'. "#EC NOTEXT
    constants c_format_date_yymmddminus type zif_excel_data_decl=>zexcel_number_format value 'yy-mm-dd'. "#EC NOTEXT
    constants c_format_date_yymmddslash type zif_excel_data_decl=>zexcel_number_format value 'yy/mm/dd'. "#EC NOTEXT
    constants c_format_date_yyyymmdd type zif_excel_data_decl=>zexcel_number_format value 'yyyymmdd'. "#EC NOTEXT
    constants c_format_date_yyyymmddminus type zif_excel_data_decl=>zexcel_number_format value 'yyyy-mm-dd'. "#EC NOTEXT
    constants c_format_date_yyyymmddslash type zif_excel_data_decl=>zexcel_number_format value 'yyyy/mm/dd'. "#EC NOTEXT
    constants c_format_date_xlsx45 type zif_excel_data_decl=>zexcel_number_format value 'mm:ss'. "#EC NOTEXT
    constants c_format_date_xlsx46 type zif_excel_data_decl=>zexcel_number_format value '[h]:mm:ss'. "#EC NOTEXT
    constants c_format_date_xlsx47 type zif_excel_data_decl=>zexcel_number_format value 'mm:ss.0'. "#EC NOTEXT
    constants c_format_general type zif_excel_data_decl=>zexcel_number_format value ''. "#EC NOTEXT
    constants c_format_number type zif_excel_data_decl=>zexcel_number_format value '0'. "#EC NOTEXT
    constants c_format_number_00 type zif_excel_data_decl=>zexcel_number_format value '0.00'. "#EC NOTEXT
    constants c_format_number_comma_sep0 type zif_excel_data_decl=>zexcel_number_format value '#,##0'. "#EC NOTEXT
    constants c_format_number_comma_sep1 type zif_excel_data_decl=>zexcel_number_format value '#,##0.00'. "#EC NOTEXT
    constants c_format_number_comma_sep2 type zif_excel_data_decl=>zexcel_number_format value '#,##0.00_-'. "#EC NOTEXT
    constants c_format_percentage type zif_excel_data_decl=>zexcel_number_format value '0%'. "#EC NOTEXT
    constants c_format_percentage_00 type zif_excel_data_decl=>zexcel_number_format value '0.00%'. "#EC NOTEXT
    constants c_format_text type zif_excel_data_decl=>zexcel_number_format value '@'. "#EC NOTEXT
    constants c_format_fraction_1 type zif_excel_data_decl=>zexcel_number_format value '# ?/?'. "#EC NOTEXT
    constants c_format_fraction_2 type zif_excel_data_decl=>zexcel_number_format value '# ??/??'. "#EC NOTEXT
    constants c_format_scientific type zif_excel_data_decl=>zexcel_number_format value '0.00E+00'. "#EC NOTEXT
    constants c_format_special_01 type zif_excel_data_decl=>zexcel_number_format value '##0.0E+0'. "#EC NOTEXT
    data format_code type zif_excel_data_decl=>zexcel_number_format .
    class-data mt_built_in_num_formats type t_num_formats read-only .
    constants c_format_xlsx37 type zif_excel_data_decl=>zexcel_number_format value '#,##0_);(#,##0)'. "#EC NOTEXT
    constants c_format_xlsx38 type zif_excel_data_decl=>zexcel_number_format value '#,##0_);[Red](#,##0)'. "#EC NOTEXT
    constants c_format_xlsx39 type zif_excel_data_decl=>zexcel_number_format value '#,##0.00_);(#,##0.00)'. "#EC NOTEXT
    constants c_format_xlsx40 type zif_excel_data_decl=>zexcel_number_format value '#,##0.00_);[Red](#,##0.00)'. "#EC NOTEXT
    constants c_format_xlsx41 type zif_excel_data_decl=>zexcel_number_format value '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'. "#EC NOTEXT
    constants c_format_xlsx42 type zif_excel_data_decl=>zexcel_number_format value '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'. "#EC NOTEXT
    constants c_format_xlsx43 type zif_excel_data_decl=>zexcel_number_format value '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'. "#EC NOTEXT
    constants c_format_xlsx44 type zif_excel_data_decl=>zexcel_number_format value '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'. "#EC NOTEXT
    constants c_format_currency_gbp_simple type zif_excel_data_decl=>zexcel_number_format value '[$£-809]#,##0.00'. "#EC NOTEXT
    constants c_format_currency_pln_simple type zif_excel_data_decl=>zexcel_number_format value '#,##0.00\ "zł"'. "#EC NOTEXT

    class-methods class_constructor .
    methods constructor .
    methods get_structure
      returning
        value(ep_number_format) type zif_excel_data_decl=>zexcel_s_style_numfmt .

*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  protected section.
  private section.
    class-methods add_format
      importing
        id   type string
        code type zif_excel_data_decl=>zexcel_number_format.
*"* private components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
endclass.



class zcl_excel_style_number_format implementation.

  method add_format.
    data ls_num_format like line of mt_built_in_num_formats.
    ls_num_format-id                  = id.
    create object ls_num_format-format.
    ls_num_format-format->format_code = code.
    insert ls_num_format into table mt_built_in_num_formats.
  endmethod.

  method class_constructor.

    clear mt_built_in_num_formats.

    add_format( id = '1' code = zcl_excel_style_number_format=>c_format_number ).               " '0'.
    add_format( id = '2' code = zcl_excel_style_number_format=>c_format_number_00 ).            " '0.00'.
    add_format( id = '3' code = zcl_excel_style_number_format=>c_format_number_comma_sep0 ).    " '#,##0'.
    add_format( id = '4' code = zcl_excel_style_number_format=>c_format_number_comma_sep1 ).    " '#,##0.00'.
    add_format( id = '5' code = zcl_excel_style_number_format=>c_format_currency_simple ).      " '$#,##0_);($#,##0)'.
    add_format( id = '6' code = zcl_excel_style_number_format=>c_format_currency_simple_red ).  " '$#,##0_);[Red]($#,##0)'.
    add_format( id = '7' code = zcl_excel_style_number_format=>c_format_currency_simple2 ).     " '$#,##0.00_);($#,##0.00)'.
    add_format( id = '8' code = zcl_excel_style_number_format=>c_format_currency_simple_red2 ). " '$#,##0.00_);[Red]($#,##0.00)'.
    add_format( id = '9' code = zcl_excel_style_number_format=>c_format_percentage ).           " '0%'.
    add_format( id = '10' code = zcl_excel_style_number_format=>c_format_percentage_00 ).        " '0.00%'.
    add_format( id = '11' code = zcl_excel_style_number_format=>c_format_scientific ).           " '0.00E+00'.
    add_format( id = '12' code = zcl_excel_style_number_format=>c_format_fraction_1 ).           " '# ?/?'.
    add_format( id = '13' code = zcl_excel_style_number_format=>c_format_fraction_2 ).           " '# ??/??'.
    add_format( id = '14' code = zcl_excel_style_number_format=>c_format_date_xlsx14 ).          "'m/d/yyyy'.  <--  should have been 'mm-dd-yy' like constant in zcl_excel_style_number_format
    add_format( id = '15' code = zcl_excel_style_number_format=>c_format_date_xlsx15 ).          "'d-mmm-yy'.
    add_format( id = '16' code = zcl_excel_style_number_format=>c_format_date_xlsx16 ).          "'d-mmm'.
    add_format( id = '17' code = zcl_excel_style_number_format=>c_format_date_xlsx17 ).          "'mmm-yy'.
    add_format( id = '18' code = zcl_excel_style_number_format=>c_format_date_time1 ).           " 'h:mm AM/PM'.
    add_format( id = '19' code = zcl_excel_style_number_format=>c_format_date_time2 ).           " 'h:mm:ss AM/PM'.
    add_format( id = '20' code = zcl_excel_style_number_format=>c_format_date_time3 ).           " 'h:mm'.
    add_format( id = '21' code = zcl_excel_style_number_format=>c_format_date_time4 ).           " 'h:mm:ss'.
    add_format( id = '22' code = zcl_excel_style_number_format=>c_format_date_xlsx22 ).          " 'm/d/yyyy h:mm'.


    add_format( id = '37' code = zcl_excel_style_number_format=>c_format_xlsx37 ).               " '#,##0_);(#,##0)'.
    add_format( id = '38' code = zcl_excel_style_number_format=>c_format_xlsx38 ).               " '#,##0_);[Red](#,##0)'.
    add_format( id = '39' code = zcl_excel_style_number_format=>c_format_xlsx39 ).               " '#,##0.00_);(#,##0.00)'.
    add_format( id = '40' code = zcl_excel_style_number_format=>c_format_xlsx40 ).               " '#,##0.00_);[Red](#,##0.00)'.
    add_format( id = '41' code = zcl_excel_style_number_format=>c_format_xlsx41 ).               " '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'.
    add_format( id = '42' code = zcl_excel_style_number_format=>c_format_xlsx42 ).               " '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'.
    add_format( id = '43' code = zcl_excel_style_number_format=>c_format_xlsx43 ).               " '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'.
    add_format( id = '44' code = zcl_excel_style_number_format=>c_format_xlsx44 ).               " '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'.
    add_format( id = '45' code = zcl_excel_style_number_format=>c_format_date_xlsx45 ).          " 'mm:ss'.
    add_format( id = '46' code = zcl_excel_style_number_format=>c_format_date_xlsx46 ).          " '[h]:mm:ss'.
    add_format( id = '47' code = zcl_excel_style_number_format=>c_format_date_xlsx47 ).          "  'mm:ss.0'.
    add_format( id = '48' code = zcl_excel_style_number_format=>c_format_special_01 ).           " '##0.0E+0'.
    add_format( id = '49' code = zcl_excel_style_number_format=>c_format_text ).                 " '@'.

  endmethod.


  method constructor.
    format_code = me->c_format_general.
  endmethod.


  method get_structure.
    ep_number_format-numfmt = me->format_code.
  endmethod.
endclass.

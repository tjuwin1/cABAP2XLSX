class zcl_excel_font definition
  public
  final
  create public .

  public section.

    types ty_font_height type n length 3.
    constants lc_default_font_height type ty_font_height value '110' ##NO_TEXT.
    constants lc_default_font_name type zif_excel_data_decl=>zexcel_style_font_name value 'Calibri' ##NO_TEXT.

    class-methods calculate_text_width
      importing
        !iv_font_name   type zif_excel_data_decl=>zexcel_style_font_name
        !iv_font_height type ty_font_height
        !iv_flag_bold   type abap_bool
        !iv_flag_italic type abap_bool
        !iv_cell_value  type zif_excel_data_decl=>zexcel_cell_value
      returning
        value(rv_width) type f .
  protected section.
  private section.
    types:
      begin of mty_s_font_metric,
        char       type c length 1,
        char_width type i,
      end of mty_s_font_metric .
    types:
      mty_th_font_metrics
               type hashed table of mty_s_font_metric
               with unique key char .
    types:
      begin of mty_s_font_cache,
        font_name       type zif_excel_data_decl=>zexcel_style_font_name,
        font_height     type ty_font_height,
        flag_bold       type abap_bool,
        flag_italic     type abap_bool,
        th_font_metrics type mty_th_font_metrics,
      end of mty_s_font_cache .
    types:
      mty_th_font_cache
               type hashed table of mty_s_font_cache
               with unique key font_name font_height flag_bold flag_italic .

    class-data mth_font_cache type mty_th_font_cache .

endclass.



class zcl_excel_font implementation.


  method calculate_text_width.
    constants lc_excel_cell_padding type f value '0.75'.
    data:         ld_length                  type i.
    ld_length = strlen( iv_cell_value ).
    rv_width = ld_length * iv_font_height / lc_default_font_height + lc_excel_cell_padding.
    if rv_width < 12.
      rv_width = 12.
    endif.

*
*
*    DATA: ld_current_character       TYPE c LENGTH 1,
*          lt_itcfc                   TYPE STANDARD TABLE OF itcfc,
*          ld_offset                  TYPE i,
*          ld_uccp                    TYPE i,
*          ls_font_metric             TYPE mty_s_font_metric,
*          ld_width_from_font_metrics TYPE i,
*          ld_font_family             TYPE itcfh-tdfamily,
*          lt_font_families           LIKE STANDARD TABLE OF ld_font_family,
*          ls_font_cache              TYPE mty_s_font_cache.
*
*    FIELD-SYMBOLS: <ls_font_cache>  TYPE mty_s_font_cache,
*                   <ls_font_metric> TYPE mty_s_font_metric,
*                   <ls_itcfc>       TYPE itcfc.
*
*    " Check if the same font (font name and font attributes) was already
*    " used before
*    READ TABLE mth_font_cache
*      WITH TABLE KEY
*        font_name   = iv_font_name
*        font_height = iv_font_height
*        flag_bold   = iv_flag_bold
*        flag_italic = iv_flag_italic
*      ASSIGNING <ls_font_cache>.
*
*    IF sy-subrc <> 0.
*      " Font is used for the first time
*      " Add the font to our local font cache
*      ls_font_cache-font_name   = iv_font_name.
*      ls_font_cache-font_height = iv_font_height.
*      ls_font_cache-flag_bold   = iv_flag_bold.
*      ls_font_cache-flag_italic = iv_flag_italic.
*      INSERT ls_font_cache INTO TABLE mth_font_cache
*        ASSIGNING <ls_font_cache>.
*
*      " Determine the SAPscript font family name from the Excel
*      " font name
**      SELECT tdfamily
**        FROM tfo01
**        WHERE tdtext = @iv_font_name
**        UP TO 1 ROWS
**        ORDER BY PRIMARY KEY
**        INTO TABLE @lt_font_families.
*
*      " Check if a matching font family was found
*      " Fonts can be uploaded from TTF files using transaction SE73
**      IF lines( lt_font_families ) > 0.
**        READ TABLE lt_font_families INDEX 1 INTO ld_font_family.
**
**        " Load font metrics (returns a table with the size of each letter
**        " in the font)
**        CALL FUNCTION 'LOAD_FONT'
**          EXPORTING
**            family      = ld_font_family
**            height      = iv_font_height
**            printer     = 'SWIN'
**            bold        = iv_flag_bold
**            italic      = iv_flag_italic
**          TABLES
**            metric      = lt_itcfc
**          EXCEPTIONS
**            font_family = 1
**            codepage    = 2
**            device_type = 3
**            OTHERS      = 4.
**        IF sy-subrc <> 0.
**          CLEAR lt_itcfc.
**        ENDIF.
**
**        " For faster access, convert each character number to the actual
**        " character, and store the characters and their sizes in a hash
**        " table
**        LOOP AT lt_itcfc ASSIGNING <ls_itcfc>.
**          ld_uccp = <ls_itcfc>-cpcharno.
**          ls_font_metric-char =
**            cl_abap_conv_in_ce=>uccpi( ld_uccp ).
**          ls_font_metric-char_width = <ls_itcfc>-tdcwidths.
**          INSERT ls_font_metric
**            INTO TABLE <ls_font_cache>-th_font_metrics.
**        ENDLOOP.
**
**      ENDIF.
**    ENDIF.
*
*    " Calculate the cell width
*    " If available, use font metrics
*    IF lines( <ls_font_cache>-th_font_metrics ) = 0.
*      " Font metrics are not available
*      " -> Calculate the cell width using only the font size
*      ld_length = strlen( iv_cell_value ).
*      rv_width = ld_length * iv_font_height / lc_default_font_height + lc_excel_cell_padding.
*
*    ELSE.
*      " Font metrics are available
*
*      " Calculate the size of the text by adding the sizes of each
*      " letter
*      ld_length = strlen( iv_cell_value ).
*      DO ld_length TIMES.
*        " Subtract 1, because the first character is at offset 0
*        ld_offset = sy-index - 1.
*
*        " Read the current character from the cell value
*        ld_current_character = iv_cell_value+ld_offset(1).
*
*        " Look up the size of the current letter
*        READ TABLE <ls_font_cache>-th_font_metrics
*          WITH TABLE KEY char = ld_current_character
*          ASSIGNING <ls_font_metric>.
*        IF sy-subrc = 0.
*          " The size of the letter is known
*          " -> Add the actual size of the letter
*          ADD <ls_font_metric>-char_width TO ld_width_from_font_metrics.
*        ELSE.
*          " The size of the letter is unknown
*          " -> Add the font height as the default letter size
*          ADD iv_font_height TO ld_width_from_font_metrics.
*        ENDIF.
*      ENDDO.
*
*      " Add cell padding (Excel makes columns a bit wider than the space
*      " that is needed for the text itself) and convert unit
*      " (division by 100)
*      rv_width = ld_width_from_font_metrics / 100 + lc_excel_cell_padding.
*    ENDIF.

  endmethod.
endclass.

class zcl_excel_style_changer definition
  public
  final
  create public .

  public section.

    interfaces zif_excel_style_changer.

    class-methods create
      importing
        excel         type ref to zcl_excel
      returning
        value(result) type ref to zif_excel_style_changer
      raising
        zcx_excel.

  protected section.
  private section.

    methods clear_initial_colorxfields
      importing
        is_color  type zif_excel_data_decl=>zexcel_s_style_color
      changing
        cs_xcolor type zif_excel_data_decl=>zexcel_s_cstylex_color.

    methods move_supplied_borders
      importing
        iv_border_supplied        type abap_bool
        is_border                 type zif_excel_data_decl=>zexcel_s_cstyle_border
        iv_xborder_supplied       type abap_bool
        is_xborder                type zif_excel_data_decl=>zexcel_s_cstylex_border
      changing
        cs_complete_style_border  type zif_excel_data_decl=>zexcel_s_cstyle_border
        cs_complete_stylex_border type zif_excel_data_decl=>zexcel_s_cstylex_border.

    data: excel                   type ref to zcl_excel,
          lv_xborder_supplied     type abap_bool,
          single_change_requested type zif_excel_data_decl=>zexcel_s_cstylex_complete,
          begin of multiple_change_requested,
            complete   type abap_bool,
            font       type abap_bool,
            fill       type abap_bool,
            begin of borders,
              complete   type abap_bool,
              allborders type abap_bool,
              diagonal   type abap_bool,
              down       type abap_bool,
              left       type abap_bool,
              right      type abap_bool,
              top        type abap_bool,
            end of borders,
            alignment  type abap_bool,
            protection type abap_bool,
          end of multiple_change_requested.
    constants:
          lv_border_supplied  type abap_bool value abap_true.
    aliases:
          complete_style   for zif_excel_style_changer~complete_style,
          complete_stylex  for zif_excel_style_changer~complete_stylex.

endclass.



class zcl_excel_style_changer implementation.


  method clear_initial_colorxfields.

    if is_color-rgb is initial.
      clear cs_xcolor-rgb.
    endif.
    if is_color-indexed is initial.
      clear cs_xcolor-indexed.
    endif.
    if is_color-theme is initial.
      clear cs_xcolor-theme.
    endif.
    if is_color-tint is initial.
      clear cs_xcolor-tint.
    endif.

  endmethod.


  method create.

    data: style type ref to zcl_excel_style_changer.

    create object style.
    style->excel = excel.
    result = style.

  endmethod.


  method move_supplied_borders.

    data: cs_borderx type zif_excel_data_decl=>zexcel_s_cstylex_border.

    if iv_border_supplied = abap_true.  " only act if parameter was supplied
      if iv_xborder_supplied = abap_true. "
        cs_borderx = is_xborder.             " use supplied x-parameter
      else.
        clear cs_borderx with 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        if is_border-border_style is  initial.
          clear cs_borderx-border_style.
        endif.
        clear_initial_colorxfields(
          exporting
            is_color  = is_border-border_color
          changing
            cs_xcolor = cs_borderx-border_color ).
      endif.
      cs_complete_style_border = is_border.
      cs_complete_stylex_border = cs_borderx.
    endif.

  endmethod.


  method zif_excel_style_changer~apply.

    data: stylemapping    type zif_excel_data_decl=>zexcel_s_stylemapping,
          lo_worksheet    type ref to zcl_excel_worksheet,
          ls_defaultcolor type zif_excel_data_decl=>zexcel_s_cstylex_color,
          l_guid          type zif_excel_data_decl=>zexcel_cell_style.
    ls_defaultcolor-rgb = abap_true.

    lo_worksheet = excel->get_worksheet_by_name( ip_sheet_name = ip_worksheet->get_title( ) ).
    if lo_worksheet <> ip_worksheet.
      zcx_excel=>raise_text( 'Worksheet doesn''t correspond to workbook of style changer'(001) ).
    endif.

    try.
        ip_worksheet->get_cell( exporting ip_column = ip_column
                                          ip_row    = ip_row
                                importing ep_guid   = l_guid ).
        stylemapping = excel->get_style_to_guid( l_guid ).
      catch zcx_excel.
* Error --> use submitted style
    endtry.


    if multiple_change_requested-complete = abap_true.
      stylemapping-complete_style = complete_style.
      stylemapping-complete_stylex = complete_stylex.
    endif.

    if multiple_change_requested-font = abap_true.
      stylemapping-complete_style-font = complete_style-font.
      stylemapping-complete_stylex-font = complete_stylex-font.
    endif.

    if multiple_change_requested-fill = abap_true.
      stylemapping-complete_style-fill = complete_style-fill.
      stylemapping-complete_stylex-fill = complete_stylex-fill.
    endif.

    if multiple_change_requested-borders-complete = abap_true.
      stylemapping-complete_style-borders = complete_style-borders.
      stylemapping-complete_stylex-borders = complete_stylex-borders.
    endif.

    if multiple_change_requested-borders-allborders = abap_true.
      stylemapping-complete_style-borders-allborders = complete_style-borders-allborders.
      stylemapping-complete_stylex-borders-allborders = complete_stylex-borders-allborders.
    endif.

    if multiple_change_requested-borders-diagonal = abap_true.
      stylemapping-complete_style-borders-diagonal = complete_style-borders-diagonal.
      stylemapping-complete_stylex-borders-diagonal = complete_stylex-borders-diagonal.
    endif.

    if multiple_change_requested-borders-down = abap_true.
      stylemapping-complete_style-borders-down = complete_style-borders-down.
      stylemapping-complete_stylex-borders-down = complete_stylex-borders-down.
    endif.

    if multiple_change_requested-borders-left = abap_true.
      stylemapping-complete_style-borders-left = complete_style-borders-left.
      stylemapping-complete_stylex-borders-left = complete_stylex-borders-left.
    endif.

    if multiple_change_requested-borders-right = abap_true.
      stylemapping-complete_style-borders-right = complete_style-borders-right.
      stylemapping-complete_stylex-borders-right = complete_stylex-borders-right.
    endif.

    if multiple_change_requested-borders-top = abap_true.
      stylemapping-complete_style-borders-top = complete_style-borders-top.
      stylemapping-complete_stylex-borders-top = complete_stylex-borders-top.
    endif.

    if multiple_change_requested-alignment = abap_true.
      stylemapping-complete_style-alignment = complete_style-alignment.
      stylemapping-complete_stylex-alignment = complete_stylex-alignment.
    endif.

    if multiple_change_requested-protection = abap_true.
      stylemapping-complete_style-protection = complete_style-protection.
      stylemapping-complete_stylex-protection = complete_stylex-protection.
    endif.

    if complete_stylex-number_format-format_code = abap_true.
      stylemapping-complete_style-number_format-format_code = complete_style-number_format-format_code.
      stylemapping-complete_stylex-number_format-format_code = abap_true.
    endif.
    if complete_stylex-font-bold = abap_true.
      stylemapping-complete_style-font-bold = complete_style-font-bold.
      stylemapping-complete_stylex-font-bold = complete_stylex-font-bold.
    endif.
    if complete_stylex-font-color = ls_defaultcolor.
      stylemapping-complete_style-font-color = complete_style-font-color.
      stylemapping-complete_stylex-font-color = complete_stylex-font-color.
    endif.
    if complete_stylex-font-color-rgb = abap_true.
      stylemapping-complete_style-font-color-rgb = complete_style-font-color-rgb.
      stylemapping-complete_stylex-font-color-rgb = complete_stylex-font-color-rgb.
    endif.
    if complete_stylex-font-color-indexed = abap_true.
      stylemapping-complete_style-font-color-indexed = complete_style-font-color-indexed.
      stylemapping-complete_stylex-font-color-indexed = complete_stylex-font-color-indexed.
    endif.
    if complete_stylex-font-color-theme = abap_true.
      stylemapping-complete_style-font-color-theme = complete_style-font-color-theme.
      stylemapping-complete_stylex-font-color-theme = complete_stylex-font-color-theme.
    endif.
    if complete_stylex-font-color-tint = abap_true.
      stylemapping-complete_style-font-color-tint = complete_style-font-color-tint.
      stylemapping-complete_stylex-font-color-tint = complete_stylex-font-color-tint.
    endif.
    if complete_stylex-font-family = abap_true.
      stylemapping-complete_style-font-family = complete_style-font-family.
      stylemapping-complete_stylex-font-family = complete_stylex-font-family.
    endif.
    if complete_stylex-font-italic = abap_true.
      stylemapping-complete_style-font-italic = complete_style-font-italic.
      stylemapping-complete_stylex-font-italic = complete_stylex-font-italic.
    endif.
    if complete_stylex-font-name = abap_true.
      stylemapping-complete_style-font-name = complete_style-font-name.
      stylemapping-complete_stylex-font-name = complete_stylex-font-name.
    endif.
    if complete_stylex-font-scheme = abap_true.
      stylemapping-complete_style-font-scheme = complete_style-font-scheme.
      stylemapping-complete_stylex-font-scheme = complete_stylex-font-scheme.
    endif.
    if complete_stylex-font-size = abap_true.
      stylemapping-complete_style-font-size = complete_style-font-size.
      stylemapping-complete_stylex-font-size = complete_stylex-font-size.
    endif.
    if complete_stylex-font-strikethrough = abap_true.
      stylemapping-complete_style-font-strikethrough = complete_style-font-strikethrough.
      stylemapping-complete_stylex-font-strikethrough = complete_stylex-font-strikethrough.
    endif.
    if complete_stylex-font-underline = abap_true.
      stylemapping-complete_style-font-underline = complete_style-font-underline.
      stylemapping-complete_stylex-font-underline = complete_stylex-font-underline.
    endif.
    if complete_stylex-font-underline_mode = abap_true.
      stylemapping-complete_style-font-underline_mode = complete_style-font-underline_mode.
      stylemapping-complete_stylex-font-underline_mode = complete_stylex-font-underline_mode.
    endif.

    if complete_stylex-fill-filltype = abap_true.
      stylemapping-complete_style-fill-filltype = complete_style-fill-filltype.
      stylemapping-complete_stylex-fill-filltype = complete_stylex-fill-filltype.
    endif.
    if complete_stylex-fill-rotation = abap_true.
      stylemapping-complete_style-fill-rotation = complete_style-fill-rotation.
      stylemapping-complete_stylex-fill-rotation = complete_stylex-fill-rotation.
    endif.
    if complete_stylex-fill-fgcolor = ls_defaultcolor.
      stylemapping-complete_style-fill-fgcolor = complete_style-fill-fgcolor.
      stylemapping-complete_stylex-fill-fgcolor = complete_stylex-fill-fgcolor.
    endif.
    if complete_stylex-fill-fgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-fgcolor-rgb = complete_style-fill-fgcolor-rgb.
      stylemapping-complete_stylex-fill-fgcolor-rgb = complete_stylex-fill-fgcolor-rgb.
    endif.
    if complete_stylex-fill-fgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-fgcolor-indexed = complete_style-fill-fgcolor-indexed.
      stylemapping-complete_stylex-fill-fgcolor-indexed = complete_stylex-fill-fgcolor-indexed.
    endif.
    if complete_stylex-fill-fgcolor-theme = abap_true.
      stylemapping-complete_style-fill-fgcolor-theme = complete_style-fill-fgcolor-theme.
      stylemapping-complete_stylex-fill-fgcolor-theme = complete_stylex-fill-fgcolor-theme.
    endif.
    if complete_stylex-fill-fgcolor-tint = abap_true.
      stylemapping-complete_style-fill-fgcolor-tint = complete_style-fill-fgcolor-tint.
      stylemapping-complete_stylex-fill-fgcolor-tint = complete_stylex-fill-fgcolor-tint.
    endif.

    if complete_stylex-fill-bgcolor = ls_defaultcolor.
      stylemapping-complete_style-fill-bgcolor = complete_style-fill-bgcolor.
      stylemapping-complete_stylex-fill-bgcolor = complete_stylex-fill-bgcolor.
    endif.
    if complete_stylex-fill-bgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-bgcolor-rgb = complete_style-fill-bgcolor-rgb.
      stylemapping-complete_stylex-fill-bgcolor-rgb = complete_stylex-fill-bgcolor-rgb.
    endif.
    if complete_stylex-fill-bgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-bgcolor-indexed = complete_style-fill-bgcolor-indexed.
      stylemapping-complete_stylex-fill-bgcolor-indexed = complete_stylex-fill-bgcolor-indexed.
    endif.
    if complete_stylex-fill-bgcolor-theme = abap_true.
      stylemapping-complete_style-fill-bgcolor-theme = complete_style-fill-bgcolor-theme.
      stylemapping-complete_stylex-fill-bgcolor-theme = complete_stylex-fill-bgcolor-theme.
    endif.
    if complete_stylex-fill-bgcolor-tint = abap_true.
      stylemapping-complete_style-fill-bgcolor-tint = complete_style-fill-bgcolor-tint.
      stylemapping-complete_stylex-fill-bgcolor-tint = complete_stylex-fill-bgcolor-tint.
    endif.

    if complete_stylex-fill-gradtype-type = abap_true.
      stylemapping-complete_style-fill-gradtype-type = complete_style-fill-gradtype-type.
      stylemapping-complete_stylex-fill-gradtype-type = complete_stylex-fill-gradtype-type.
    endif.
    if complete_stylex-fill-gradtype-degree = abap_true.
      stylemapping-complete_style-fill-gradtype-degree = complete_style-fill-gradtype-degree.
      stylemapping-complete_stylex-fill-gradtype-degree = complete_stylex-fill-gradtype-degree.
    endif.
    if complete_stylex-fill-gradtype-bottom = abap_true.
      stylemapping-complete_style-fill-gradtype-bottom = complete_style-fill-gradtype-bottom.
      stylemapping-complete_stylex-fill-gradtype-bottom = complete_stylex-fill-gradtype-bottom.
    endif.
    if complete_stylex-fill-gradtype-left = abap_true.
      stylemapping-complete_style-fill-gradtype-left = complete_style-fill-gradtype-left.
      stylemapping-complete_stylex-fill-gradtype-left = complete_stylex-fill-gradtype-left.
    endif.
    if complete_stylex-fill-gradtype-top = abap_true.
      stylemapping-complete_style-fill-gradtype-top = complete_style-fill-gradtype-top.
      stylemapping-complete_stylex-fill-gradtype-top = complete_stylex-fill-gradtype-top.
    endif.
    if complete_stylex-fill-gradtype-right = abap_true.
      stylemapping-complete_style-fill-gradtype-right = complete_style-fill-gradtype-right.
      stylemapping-complete_stylex-fill-gradtype-right = complete_stylex-fill-gradtype-right.
    endif.
    if complete_stylex-fill-gradtype-position1 = abap_true.
      stylemapping-complete_style-fill-gradtype-position1 = complete_style-fill-gradtype-position1.
      stylemapping-complete_stylex-fill-gradtype-position1 = complete_stylex-fill-gradtype-position1.
    endif.
    if complete_stylex-fill-gradtype-position2 = abap_true.
      stylemapping-complete_style-fill-gradtype-position2 = complete_style-fill-gradtype-position2.
      stylemapping-complete_stylex-fill-gradtype-position2 = complete_stylex-fill-gradtype-position2.
    endif.
    if complete_stylex-fill-gradtype-position3 = abap_true.
      stylemapping-complete_style-fill-gradtype-position3 = complete_style-fill-gradtype-position3.
      stylemapping-complete_stylex-fill-gradtype-position3 = complete_stylex-fill-gradtype-position3.
    endif.



    if complete_stylex-borders-diagonal_mode = abap_true.
      stylemapping-complete_style-borders-diagonal_mode = complete_style-borders-diagonal_mode.
      stylemapping-complete_stylex-borders-diagonal_mode = complete_stylex-borders-diagonal_mode.
    endif.
    if complete_stylex-alignment-horizontal = abap_true.
      stylemapping-complete_style-alignment-horizontal = complete_style-alignment-horizontal.
      stylemapping-complete_stylex-alignment-horizontal = complete_stylex-alignment-horizontal.
    endif.
    if complete_stylex-alignment-vertical = abap_true.
      stylemapping-complete_style-alignment-vertical = complete_style-alignment-vertical.
      stylemapping-complete_stylex-alignment-vertical = complete_stylex-alignment-vertical.
    endif.
    if complete_stylex-alignment-textrotation = abap_true.
      stylemapping-complete_style-alignment-textrotation = complete_style-alignment-textrotation.
      stylemapping-complete_stylex-alignment-textrotation = complete_stylex-alignment-textrotation.
    endif.
    if complete_stylex-alignment-wraptext = abap_true.
      stylemapping-complete_style-alignment-wraptext = complete_style-alignment-wraptext.
      stylemapping-complete_stylex-alignment-wraptext = complete_stylex-alignment-wraptext.
    endif.
    if complete_stylex-alignment-shrinktofit = abap_true.
      stylemapping-complete_style-alignment-shrinktofit = complete_style-alignment-shrinktofit.
      stylemapping-complete_stylex-alignment-shrinktofit = complete_stylex-alignment-shrinktofit.
    endif.
    if complete_stylex-alignment-indent = abap_true.
      stylemapping-complete_style-alignment-indent = complete_style-alignment-indent.
      stylemapping-complete_stylex-alignment-indent = complete_stylex-alignment-indent.
    endif.
    if complete_stylex-protection-hidden = abap_true.
      stylemapping-complete_style-protection-hidden = complete_style-protection-hidden.
      stylemapping-complete_stylex-protection-hidden = complete_stylex-protection-hidden.
    endif.
    if complete_stylex-protection-locked = abap_true.
      stylemapping-complete_style-protection-locked = complete_style-protection-locked.
      stylemapping-complete_stylex-protection-locked = complete_stylex-protection-locked.
    endif.

    if complete_stylex-borders-allborders-border_style = abap_true.
      stylemapping-complete_style-borders-allborders-border_style = complete_style-borders-allborders-border_style.
      stylemapping-complete_stylex-borders-allborders-border_style = complete_stylex-borders-allborders-border_style.
    endif.
    if complete_stylex-borders-allborders-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-rgb = complete_style-borders-allborders-border_color-rgb.
      stylemapping-complete_stylex-borders-allborders-border_color-rgb = complete_stylex-borders-allborders-border_color-rgb.
    endif.
    if complete_stylex-borders-allborders-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-indexed = complete_style-borders-allborders-border_color-indexed.
      stylemapping-complete_stylex-borders-allborders-border_color-indexed = complete_stylex-borders-allborders-border_color-indexed.
    endif.
    if complete_stylex-borders-allborders-border_color-theme = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-theme = complete_style-borders-allborders-border_color-theme.
      stylemapping-complete_stylex-borders-allborders-border_color-theme = complete_stylex-borders-allborders-border_color-theme.
    endif.
    if complete_stylex-borders-allborders-border_color-tint = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-tint = complete_style-borders-allborders-border_color-tint.
      stylemapping-complete_stylex-borders-allborders-border_color-tint = complete_stylex-borders-allborders-border_color-tint.
    endif.

    if complete_stylex-borders-diagonal-border_style = abap_true.
      stylemapping-complete_style-borders-diagonal-border_style = complete_style-borders-diagonal-border_style.
      stylemapping-complete_stylex-borders-diagonal-border_style = complete_stylex-borders-diagonal-border_style.
    endif.
    if complete_stylex-borders-diagonal-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-rgb = complete_style-borders-diagonal-border_color-rgb.
      stylemapping-complete_stylex-borders-diagonal-border_color-rgb = complete_stylex-borders-diagonal-border_color-rgb.
    endif.
    if complete_stylex-borders-diagonal-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-indexed = complete_style-borders-diagonal-border_color-indexed.
      stylemapping-complete_stylex-borders-diagonal-border_color-indexed = complete_stylex-borders-diagonal-border_color-indexed.
    endif.
    if complete_stylex-borders-diagonal-border_color-theme = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-theme = complete_style-borders-diagonal-border_color-theme.
      stylemapping-complete_stylex-borders-diagonal-border_color-theme = complete_stylex-borders-diagonal-border_color-theme.
    endif.
    if complete_stylex-borders-diagonal-border_color-tint = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-tint = complete_style-borders-diagonal-border_color-tint.
      stylemapping-complete_stylex-borders-diagonal-border_color-tint = complete_stylex-borders-diagonal-border_color-tint.
    endif.

    if complete_stylex-borders-down-border_style = abap_true.
      stylemapping-complete_style-borders-down-border_style = complete_style-borders-down-border_style.
      stylemapping-complete_stylex-borders-down-border_style = complete_stylex-borders-down-border_style.
    endif.
    if complete_stylex-borders-down-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-down-border_color-rgb = complete_style-borders-down-border_color-rgb.
      stylemapping-complete_stylex-borders-down-border_color-rgb = complete_stylex-borders-down-border_color-rgb.
    endif.
    if complete_stylex-borders-down-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-down-border_color-indexed = complete_style-borders-down-border_color-indexed.
      stylemapping-complete_stylex-borders-down-border_color-indexed = complete_stylex-borders-down-border_color-indexed.
    endif.
    if complete_stylex-borders-down-border_color-theme = abap_true.
      stylemapping-complete_style-borders-down-border_color-theme = complete_style-borders-down-border_color-theme.
      stylemapping-complete_stylex-borders-down-border_color-theme = complete_stylex-borders-down-border_color-theme.
    endif.
    if complete_stylex-borders-down-border_color-tint = abap_true.
      stylemapping-complete_style-borders-down-border_color-tint = complete_style-borders-down-border_color-tint.
      stylemapping-complete_stylex-borders-down-border_color-tint = complete_stylex-borders-down-border_color-tint.
    endif.

    if complete_stylex-borders-left-border_style = abap_true.
      stylemapping-complete_style-borders-left-border_style = complete_style-borders-left-border_style.
      stylemapping-complete_stylex-borders-left-border_style = complete_stylex-borders-left-border_style.
    endif.
    if complete_stylex-borders-left-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-left-border_color-rgb = complete_style-borders-left-border_color-rgb.
      stylemapping-complete_stylex-borders-left-border_color-rgb = complete_stylex-borders-left-border_color-rgb.
    endif.
    if complete_stylex-borders-left-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-left-border_color-indexed = complete_style-borders-left-border_color-indexed.
      stylemapping-complete_stylex-borders-left-border_color-indexed = complete_stylex-borders-left-border_color-indexed.
    endif.
    if complete_stylex-borders-left-border_color-theme = abap_true.
      stylemapping-complete_style-borders-left-border_color-theme = complete_style-borders-left-border_color-theme.
      stylemapping-complete_stylex-borders-left-border_color-theme = complete_stylex-borders-left-border_color-theme.
    endif.
    if complete_stylex-borders-left-border_color-tint = abap_true.
      stylemapping-complete_style-borders-left-border_color-tint = complete_style-borders-left-border_color-tint.
      stylemapping-complete_stylex-borders-left-border_color-tint = complete_stylex-borders-left-border_color-tint.
    endif.

    if complete_stylex-borders-right-border_style = abap_true.
      stylemapping-complete_style-borders-right-border_style = complete_style-borders-right-border_style.
      stylemapping-complete_stylex-borders-right-border_style = complete_stylex-borders-right-border_style.
    endif.
    if complete_stylex-borders-right-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-right-border_color-rgb = complete_style-borders-right-border_color-rgb.
      stylemapping-complete_stylex-borders-right-border_color-rgb = complete_stylex-borders-right-border_color-rgb.
    endif.
    if complete_stylex-borders-right-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-right-border_color-indexed = complete_style-borders-right-border_color-indexed.
      stylemapping-complete_stylex-borders-right-border_color-indexed = complete_stylex-borders-right-border_color-indexed.
    endif.
    if complete_stylex-borders-right-border_color-theme = abap_true.
      stylemapping-complete_style-borders-right-border_color-theme = complete_style-borders-right-border_color-theme.
      stylemapping-complete_stylex-borders-right-border_color-theme = complete_stylex-borders-right-border_color-theme.
    endif.
    if complete_stylex-borders-right-border_color-tint = abap_true.
      stylemapping-complete_style-borders-right-border_color-tint = complete_style-borders-right-border_color-tint.
      stylemapping-complete_stylex-borders-right-border_color-tint = complete_stylex-borders-right-border_color-tint.
    endif.

    if complete_stylex-borders-top-border_style = abap_true.
      stylemapping-complete_style-borders-top-border_style = complete_style-borders-top-border_style.
      stylemapping-complete_stylex-borders-top-border_style = complete_stylex-borders-top-border_style.
    endif.
    if complete_stylex-borders-top-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-top-border_color-rgb = complete_style-borders-top-border_color-rgb.
      stylemapping-complete_stylex-borders-top-border_color-rgb = complete_stylex-borders-top-border_color-rgb.
    endif.
    if complete_stylex-borders-top-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-top-border_color-indexed = complete_style-borders-top-border_color-indexed.
      stylemapping-complete_stylex-borders-top-border_color-indexed = complete_stylex-borders-top-border_color-indexed.
    endif.
    if complete_stylex-borders-top-border_color-theme = abap_true.
      stylemapping-complete_style-borders-top-border_color-theme = complete_style-borders-top-border_color-theme.
      stylemapping-complete_stylex-borders-top-border_color-theme = complete_stylex-borders-top-border_color-theme.
    endif.
    if complete_stylex-borders-top-border_color-tint = abap_true.
      stylemapping-complete_style-borders-top-border_color-tint = complete_style-borders-top-border_color-tint.
      stylemapping-complete_stylex-borders-top-border_color-tint = complete_stylex-borders-top-border_color-tint.
    endif.


* Now we have a completly filled styles.
* This can be used to get the guid
* Return guid if requested.  Might be used if copy&paste of styles is requested
    ep_guid = me->excel->get_static_cellstyle_guid( ip_cstyle_complete  = stylemapping-complete_style
                                                   ip_cstylex_complete = stylemapping-complete_stylex  ).
    lo_worksheet->set_cell_style( ip_column = ip_column
                                  ip_row    = ip_row
                                  ip_style  = ep_guid ).

  endmethod.


  method zif_excel_style_changer~get_guid.

    result = excel->get_static_cellstyle_guid( ip_cstyle_complete  = complete_style
                                               ip_cstylex_complete = complete_stylex  ).

  endmethod.


  method zif_excel_style_changer~set_alignment_horizontal.

    complete_style-alignment-horizontal = value.
    complete_stylex-alignment-horizontal = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_alignment_indent.

    complete_style-alignment-indent = value.
    complete_stylex-alignment-indent = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_alignment_shrinktofit.

    complete_style-alignment-shrinktofit = value.
    complete_stylex-alignment-shrinktofit = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_alignment_textrotation.

    complete_style-alignment-textrotation = value.
    complete_stylex-alignment-textrotation = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_alignment_vertical.

    complete_style-alignment-vertical = value.
    complete_stylex-alignment-vertical = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_alignment_wraptext.

    complete_style-alignment-wraptext = value.
    complete_stylex-alignment-wraptext = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allborders_color.

    complete_style-borders-allborders-border_color = value.
    complete_stylex-borders-allborders-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allborders_style.

    complete_style-borders-allborders-border_style = value.
    complete_stylex-borders-allborders-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allbo_color_indexe.

    complete_style-borders-allborders-border_color-indexed = value.
    complete_stylex-borders-allborders-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allbo_color_rgb.

    complete_style-borders-allborders-border_color-rgb = value.
    complete_stylex-borders-allborders-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allbo_color_theme.

    complete_style-borders-allborders-border_color-theme = value.
    complete_stylex-borders-allborders-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_allbo_color_tint.

    complete_style-borders-allborders-border_color-tint = value.
    complete_stylex-borders-allborders-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_color.

    complete_style-borders-diagonal-border_color = value.
    complete_stylex-borders-diagonal-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_color_ind.

    complete_style-borders-diagonal-border_color-indexed = value.
    complete_stylex-borders-diagonal-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_color_rgb.

    complete_style-borders-diagonal-border_color-rgb = value.
    complete_stylex-borders-diagonal-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_color_the.

    complete_style-borders-diagonal-border_color-theme = value.
    complete_stylex-borders-diagonal-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_color_tin.

    complete_style-borders-diagonal-border_color-tint = value.
    complete_stylex-borders-diagonal-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_mode.

    complete_style-borders-diagonal_mode = value.
    complete_stylex-borders-diagonal_mode = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_diagonal_style.

    complete_style-borders-diagonal-border_style = value.
    complete_stylex-borders-diagonal-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_color.

    complete_style-borders-down-border_color = value.
    complete_stylex-borders-down-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_color_indexed.

    complete_style-borders-down-border_color-indexed = value.
    complete_stylex-borders-down-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_color_rgb.

    complete_style-borders-down-border_color-rgb = value.
    complete_stylex-borders-down-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_color_theme.

    complete_style-borders-down-border_color-theme = value.
    complete_stylex-borders-down-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_color_tint.

    complete_style-borders-down-border_color-tint = value.
    complete_stylex-borders-down-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_down_style.

    complete_style-borders-down-border_style = value.
    complete_stylex-borders-down-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_color.

    complete_style-borders-left-border_color = value.
    complete_stylex-borders-left-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_color_indexed.

    complete_style-borders-left-border_color-indexed = value.
    complete_stylex-borders-left-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_color_rgb.

    complete_style-borders-left-border_color-rgb = value.
    complete_stylex-borders-left-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_color_theme.

    complete_style-borders-left-border_color-theme = value.
    complete_stylex-borders-left-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_color_tint.

    complete_style-borders-left-border_color-tint = value.
    complete_stylex-borders-left-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_left_style.

    complete_style-borders-left-border_style = value.
    complete_stylex-borders-left-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_color.

    complete_style-borders-right-border_color = value.
    complete_stylex-borders-right-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_color_indexe.

    complete_style-borders-right-border_color-indexed = value.
    complete_stylex-borders-right-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_color_rgb.

    complete_style-borders-right-border_color-rgb = value.
    complete_stylex-borders-right-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_color_theme.

    complete_style-borders-right-border_color-theme = value.
    complete_stylex-borders-right-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_color_tint.

    complete_style-borders-right-border_color-tint = value.
    complete_stylex-borders-right-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_right_style.

    complete_style-borders-right-border_style = value.
    complete_stylex-borders-right-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_color.

    complete_style-borders-top-border_color = value.
    complete_stylex-borders-top-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_color_indexed.

    complete_style-borders-top-border_color-indexed = value.
    complete_stylex-borders-top-border_color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_color_rgb.

    complete_style-borders-top-border_color-rgb = value.
    complete_stylex-borders-top-border_color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_color_theme.

    complete_style-borders-top-border_color-theme = value.
    complete_stylex-borders-top-border_color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_color_tint.

    complete_style-borders-top-border_color-tint = value.
    complete_stylex-borders-top-border_color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_borders_top_style.

    complete_style-borders-top-border_style = value.
    complete_stylex-borders-top-border_style = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete.

    complete_style = ip_complete.
    complete_stylex = ip_xcomplete.
    multiple_change_requested-complete = abap_true.
    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_alignment.

    data: alignmentx like ip_xalignment.

    if ip_xalignment is supplied.
      alignmentx = ip_xalignment.
    else.
      clear alignmentx with 'X'.
      if ip_alignment-horizontal is initial.
        clear alignmentx-horizontal.
      endif.
      if ip_alignment-vertical is initial.
        clear alignmentx-vertical.
      endif.
    endif.

    complete_style-alignment  = ip_alignment .
    complete_stylex-alignment = alignmentx   .
    multiple_change_requested-alignment = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders.

    data: bordersx like ip_xborders.
    if ip_xborders is supplied.
      bordersx = ip_xborders.
    else.
      clear bordersx with 'X'.
      if ip_borders-allborders-border_style is initial.
        clear bordersx-allborders-border_style.
      endif.
      if ip_borders-diagonal-border_style is initial.
        clear bordersx-diagonal-border_style.
      endif.
      if ip_borders-down-border_style is initial.
        clear bordersx-down-border_style.
      endif.
      if ip_borders-left-border_style is initial.
        clear bordersx-left-border_style.
      endif.
      if ip_borders-right-border_style is initial.
        clear bordersx-right-border_style.
      endif.
      if ip_borders-top-border_style is initial.
        clear bordersx-top-border_style.
      endif.

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-allborders-border_color
        changing
          cs_xcolor = bordersx-allborders-border_color ).

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-diagonal-border_color
        changing
          cs_xcolor = bordersx-diagonal-border_color ).

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-down-border_color
        changing
          cs_xcolor = bordersx-down-border_color ).

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-left-border_color
        changing
          cs_xcolor = bordersx-left-border_color ).

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-right-border_color
        changing
          cs_xcolor = bordersx-right-border_color ).

      clear_initial_colorxfields(
        exporting
          is_color  = ip_borders-top-border_color
        changing
          cs_xcolor = bordersx-top-border_color ).

    endif.

    complete_style-borders = ip_borders.
    complete_stylex-borders = bordersx.
    multiple_change_requested-borders-complete = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_all.

    lv_xborder_supplied = boolc( ip_xborders_allborders is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_allborders
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_allborders
      changing
        cs_complete_style_border  = complete_style-borders-allborders
        cs_complete_stylex_border = complete_stylex-borders-allborders ).
    multiple_change_requested-borders-allborders = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_diagonal.

    lv_xborder_supplied = boolc( ip_xborders_diagonal is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_diagonal
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_diagonal
      changing
        cs_complete_style_border  = complete_style-borders-diagonal
        cs_complete_stylex_border = complete_stylex-borders-diagonal ).
    multiple_change_requested-borders-diagonal = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_down.

    lv_xborder_supplied = boolc( ip_xborders_down is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_down
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_down
      changing
        cs_complete_style_border  = complete_style-borders-down
        cs_complete_stylex_border = complete_stylex-borders-down ).
    multiple_change_requested-borders-down = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_left.

    lv_xborder_supplied = boolc( ip_xborders_left is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_left
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_left
      changing
        cs_complete_style_border  = complete_style-borders-left
        cs_complete_stylex_border = complete_stylex-borders-left ).
    multiple_change_requested-borders-left = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_right.

    lv_xborder_supplied = boolc( ip_xborders_right is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_right
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_right
      changing
        cs_complete_style_border  = complete_style-borders-right
        cs_complete_stylex_border = complete_stylex-borders-right ).
    multiple_change_requested-borders-right = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_borders_top.

    lv_xborder_supplied = boolc( ip_xborders_top is supplied ).
    move_supplied_borders(
      exporting
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_top
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_top
      changing
        cs_complete_style_border  = complete_style-borders-top
        cs_complete_stylex_border = complete_stylex-borders-top ).
    multiple_change_requested-borders-top = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_fill.

    data: fillx like ip_xfill.
    if ip_xfill is supplied.
      fillx = ip_xfill.
    else.
      clear fillx with 'X'.
      if ip_fill-filltype is initial.
        clear fillx-filltype.
      endif.
      clear_initial_colorxfields(
        exporting
          is_color  = ip_fill-fgcolor
        changing
          cs_xcolor = fillx-fgcolor ).
      clear_initial_colorxfields(
        exporting
          is_color  = ip_fill-bgcolor
        changing
          cs_xcolor = fillx-bgcolor ).

    endif.

    complete_style-fill = ip_fill.
    complete_stylex-fill = fillx.
    multiple_change_requested-fill = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_font.

    data: fontx type zif_excel_data_decl=>zexcel_s_cstylex_font.

    if ip_xfont is supplied.
      fontx = ip_xfont.
    else.
* Only supplied values should be used - exception: Flags bold and italic strikethrough underline
      fontx-bold = 'X'.
      fontx-italic = 'X'.
      fontx-strikethrough = 'X'.
      fontx-underline_mode = 'X'.
      clear fontx-color with 'X'.
      clear_initial_colorxfields(
        exporting
          is_color  = ip_font-color
        changing
          cs_xcolor = fontx-color ).
      if ip_font-family is not initial.
        fontx-family = 'X'.
      endif.
      if ip_font-name is not initial.
        fontx-name = 'X'.
      endif.
      if ip_font-scheme is not initial.
        fontx-scheme = 'X'.
      endif.
      if ip_font-size is not initial.
        fontx-size = 'X'.
      endif.
      if ip_font-underline_mode is not initial.
        fontx-underline_mode = 'X'.
      endif.
    endif.

    complete_style-font = ip_font.
    complete_stylex-font = fontx.
    multiple_change_requested-font = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_complete_protection.

    move-corresponding ip_protection  to complete_style-protection.
    if ip_xprotection is supplied.
      move-corresponding ip_xprotection to complete_stylex-protection.
    else.
      if ip_protection-hidden is not initial.
        complete_stylex-protection-hidden = 'X'.
      endif.
      if ip_protection-locked is not initial.
        complete_stylex-protection-locked = 'X'.
      endif.
    endif.
    multiple_change_requested-protection = abap_true.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_bgcolor.

    complete_style-fill-bgcolor = value.
    complete_stylex-fill-bgcolor-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_bgcolor_indexed.

    complete_style-fill-bgcolor-indexed = value.
    complete_stylex-fill-bgcolor-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_bgcolor_rgb.

    complete_style-fill-bgcolor-rgb = value.
    complete_stylex-fill-bgcolor-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_bgcolor_theme.

    complete_style-fill-bgcolor-theme = value.
    complete_stylex-fill-bgcolor-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_bgcolor_tint.

    complete_style-fill-bgcolor-tint = value.
    complete_stylex-fill-bgcolor-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_fgcolor.

    complete_style-fill-fgcolor = value.
    complete_stylex-fill-fgcolor-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_fgcolor_indexed.

    complete_style-fill-fgcolor-indexed = value.
    complete_stylex-fill-fgcolor-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_fgcolor_rgb.

    complete_style-fill-fgcolor-rgb = value.
    complete_stylex-fill-fgcolor-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_fgcolor_theme.

    complete_style-fill-fgcolor-theme = value.
    complete_stylex-fill-fgcolor-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_fgcolor_tint.

    complete_style-fill-fgcolor-tint = value.
    complete_stylex-fill-fgcolor-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_filltype.

    complete_style-fill-filltype = value.
    complete_stylex-fill-filltype = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_bottom.

    complete_style-fill-gradtype-bottom = value.
    complete_stylex-fill-gradtype-bottom = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_degree.

    complete_style-fill-gradtype-degree = value.
    complete_stylex-fill-gradtype-degree = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_left.

    complete_style-fill-gradtype-left = value.
    complete_stylex-fill-gradtype-left = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_position1.

    complete_style-fill-gradtype-position1 = value.
    complete_stylex-fill-gradtype-position1 = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_position2.

    complete_style-fill-gradtype-position2 = value.
    complete_stylex-fill-gradtype-position2 = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_position3.

    complete_style-fill-gradtype-position3 = value.
    complete_stylex-fill-gradtype-position3 = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_right.

    complete_style-fill-gradtype-right = value.
    complete_stylex-fill-gradtype-right = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_top.

    complete_style-fill-gradtype-top = value.
    complete_stylex-fill-gradtype-top = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_gradtype_type.

    complete_style-fill-gradtype-type = value.
    complete_stylex-fill-gradtype-type = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_fill_rotation.

    complete_style-fill-rotation = value.
    complete_stylex-fill-rotation = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_bold.

    complete_style-font-bold = value.
    complete_stylex-font-bold = 'X'.
    single_change_requested-font-bold = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_color.

    complete_style-font-color = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_color_indexed.

    complete_style-font-color-indexed = value.
    complete_stylex-font-color-indexed = 'X'.
    single_change_requested-font-color-indexed = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_color_rgb.

    complete_style-font-color-rgb = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_color_theme.

    complete_style-font-color-theme = value.
    complete_stylex-font-color-theme = 'X'.
    single_change_requested-font-color-theme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_color_tint.

    complete_style-font-color-tint = value.
    complete_stylex-font-color-tint = 'X'.
    single_change_requested-font-color-tint = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_family.

    complete_style-font-family = value.
    complete_stylex-font-family = 'X'.
    single_change_requested-font-family = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_italic.

    complete_style-font-italic = value.
    complete_stylex-font-italic = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_name.

    complete_style-font-name = value.
    complete_stylex-font-name = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_scheme.

    complete_style-font-scheme = value.
    complete_stylex-font-scheme = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_size.

    complete_style-font-size = value.
    complete_stylex-font-size = abap_true.
    single_change_requested-font-size = abap_true.
    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_strikethrough.

    complete_style-font-strikethrough = value.
    complete_stylex-font-strikethrough = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_underline.

    complete_style-font-underline = value.
    complete_stylex-font-underline = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_font_underline_mode.

    complete_style-font-underline_mode = value.
    complete_stylex-font-underline_mode = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_number_format.

    complete_style-number_format-format_code = value.
    complete_stylex-number_format-format_code = abap_true.
    single_change_requested-number_format-format_code = abap_true.
    result = me.

  endmethod.


  method zif_excel_style_changer~set_protection_hidden.

    complete_style-protection-hidden = value.
    complete_stylex-protection-hidden = 'X'.

    result = me.

  endmethod.


  method zif_excel_style_changer~set_protection_locked.

    complete_style-protection-locked = value.
    complete_stylex-protection-locked = 'X'.

    result = me.

  endmethod.
endclass.

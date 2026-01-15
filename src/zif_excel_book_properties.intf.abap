INTERFACE zif_excel_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE zif_excel_data_decl=>zexcel_creator .
  DATA lastmodifiedby TYPE zif_excel_data_decl=>zexcel_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE zif_excel_data_decl=>zexcel_title .
  DATA subject TYPE zif_excel_data_decl=>zexcel_subject .
  DATA description TYPE zif_excel_data_decl=>zexcel_description .
  DATA keywords TYPE zif_excel_data_decl=>zexcel_keywords .
  DATA category TYPE zif_excel_data_decl=>zexcel_category .
  DATA company TYPE zif_excel_data_decl=>zexcel_company .
  DATA application TYPE zif_excel_data_decl=>zexcel_application .
  DATA docsecurity TYPE zif_excel_data_decl=>zexcel_docsecurity .
  DATA scalecrop TYPE zif_excel_data_decl=>zexcel_scalecrop .
  DATA linksuptodate TYPE abap_boolean.
  DATA shareddoc TYPE abap_boolean .
  DATA hyperlinkschanged TYPE abap_boolean .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.

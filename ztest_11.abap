REPORT ztest_11.

CLASS zcl_excel DEFINITION.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_data,
             col1 TYPE string,
             col2 TYPE string,
             col3 TYPE string,
             col4 TYPE string,
           END OF ty_data.

    DATA: lt_data TYPE STANDARD TABLE OF ty_data.
    DATA: lo_ixml     TYPE REF TO if_ixml,
          lo_document TYPE REF TO if_ixml_document.
    DATA: lo_root TYPE REF TO if_ixml_element.
    DATA: lo_worksheet TYPE REF TO if_ixml_element.
    DATA: lo_table TYPE REF TO if_ixml_element.
    DATA: lo_row  TYPE REF TO if_ixml_element,
          lo_cell TYPE REF TO if_ixml_element,
          lo_data TYPE REF TO if_ixml_element.

    TYPES: BEGIN OF ty_emp,
             empid TYPE string,
             name  TYPE string,
           END OF ty_emp.

    DATA: gt_emp TYPE TABLE OF ty_emp,
          gs_emp TYPE ty_emp.

    DATA: lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lv_xml           TYPE string,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    METHODS create_xml.
    METHODS create_workbook.
    METHODS create_worksheet.
    METHODS create_table.
    METHODS create_header_row.
    METHODS add_data_rows.
    METHODS create_styles.
    METHODS render_xml.
    METHODS download_excel.
  PROTECTED SECTION.
  PRIVATE SECTION.

ENDCLASS.

CLASS zcl_excel IMPLEMENTATION.
  METHOD create_xml.
    lo_ixml = cl_ixml=>create( ).
    lo_document = lo_ixml->create_document( ).
  ENDMETHOD.
  METHOD create_workbook.
    lo_root =
      lo_document->create_simple_element(
        name = 'Workbook'
        parent = lo_document ).

    lo_root->set_attribute(
      name = 'xmlns'
      value = 'urn:schemas-microsoft-com:office:spreadsheet' ).
  ENDMETHOD.
  METHOD create_worksheet.
    lo_worksheet =
      lo_document->create_simple_element(
        name = 'Worksheet'
        parent = lo_root ).

    lo_worksheet->set_attribute(
      name = 'ss:Name'
      value = 'Sheet1' ).
  ENDMETHOD.
  METHOD create_table.
    lo_table = lo_document->create_simple_element(
    name = 'Table'
    parent = lo_worksheet ).
  ENDMETHOD.
  METHOD create_header_row.
    lo_row = lo_document->create_simple_element(
      name = 'Row'
      parent = lo_table ).

    " Column 1
    lo_cell = lo_document->create_simple_element(
      name = 'Cell'
      parent = lo_row ).

    lo_data = lo_document->create_simple_element(
      name = 'Data'
      parent = lo_cell ).

    lo_data->set_attribute(
    name = 'ss:Type'
    value = 'String' ).

    lo_data->if_ixml_node~set_value( 'Employee ID' ).

    " Column 2
    lo_cell = lo_document->create_simple_element(
      name = 'Cell'
      parent = lo_row ).

    lo_data = lo_document->create_simple_element(
      name = 'Data'
      parent = lo_cell ).

    lo_data->set_attribute(
      name = 'ss:Type'
      value = 'String' ).

    lo_data->if_ixml_node~set_value( 'Employee Name' ).
  ENDMETHOD.
  METHOD add_data_rows.

    lt_data = VALUE #( ( col1 = 'Yatchting'
                         col2 = 'Services'
                         col3 = 'Workboats'
                         col4 = 'Global Support' )
                       ( col1 = '71'
                         col2 = '39'
                         col3 = '59' )
                       ( col1 = '' )
                       ( col1 = 'FUE Factor'
                         col2 = 'Factor'   )
                       ( col1 = 'Professional Use'
                         col2 = '1' )
                       ( col1 = 'Functional Use'
                         col2 = '5' )
                       ( col1 = 'Professional Use'
                         col2 = '1' )  ).

    FIELD-SYMBOLS:
      <fs_comp> TYPE any.

    DATA:
      lr_descr TYPE REF TO cl_abap_structdescr,
      lt_comp  TYPE cl_abap_structdescr=>component_table.

    lr_descr ?= cl_abap_typedescr=>describe_by_data(
                  VALUE ty_data( ) ).

    lt_comp = lr_descr->get_components( ).

    LOOP AT lt_data INTO DATA(ls_data).

      DATA(lv_row_idx) = sy-tabix.

      lo_row = lo_document->create_simple_element(
                 name   = 'Row'
                 parent = lo_table ).

      LOOP AT lt_comp INTO DATA(ls_comp).
        ASSIGN COMPONENT ls_comp-name
               OF STRUCTURE ls_data
               TO <fs_comp>.

        IF sy-subrc = 0.

          lo_cell = lo_document->create_simple_element(
                      name   = 'Cell'
                      parent = lo_row ).
          " Apply bold style to rows 1 and 4
          IF lv_row_idx = 1 OR lv_row_idx = 4.
            lo_cell->set_attribute(
              name  = 'ss:StyleID'
              value = 'BoldStyle' ).
          ENDIF.
          lo_data = lo_document->create_simple_element(
                      name   = 'Data'
                      parent = lo_cell ).

          lo_data->set_attribute(
            name  = 'ss:Type'
            value = 'String' ).

          lo_data->if_ixml_node~set_value(
            CONV string( <fs_comp> ) ).

        ENDIF.

      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.
  METHOD create_styles.

    DATA: lo_styles TYPE REF TO if_ixml_element,
          lo_style  TYPE REF TO if_ixml_element,
          lo_font   TYPE REF TO if_ixml_element.

    lo_styles = lo_document->create_simple_element(
    name = 'Styles'
    parent = lo_root ).

    lo_style = lo_document->create_simple_element(
    name = 'Style'
    parent = lo_styles ).

    lo_style->set_attribute(
    name = 'ss:ID'
    value = 'BoldStyle' ).

    lo_font = lo_document->create_simple_element(
    name = 'Font'
    parent = lo_style ).

    lo_font->set_attribute(
    name = 'ss:Bold'
    value = '1' ).

  ENDMETHOD.
  METHOD render_xml.
    lo_streamfactory = lo_ixml->create_stream_factory( ).

    lo_ostream = lo_streamfactory->create_ostream_cstring(
    string = lv_xml ).

    lo_renderer = lo_ixml->create_renderer(
    document = lo_document
    ostream = lo_ostream ).

    lo_renderer->render( ).
  ENDMETHOD.
  METHOD download_excel.
    DATA lt_bin TYPE solix_tab.
    DATA lv_xstr TYPE xstring.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = lv_xml
      IMPORTING
        buffer = lv_xstr.

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer     = lv_xstr
      TABLES
        binary_tab = lt_bin.

    cl_gui_frontend_services=>gui_download(
    EXPORTING
    filename = 'C:\Users\Nikhil.Mishra\Documents\EmployeeReport.xls'
    filetype = 'BIN'
    CHANGING
    data_tab = lt_bin ).

  ENDMETHOD.
ENDCLASS.

START-OF-SELECTION.

  DATA(lo_data) = NEW zcl_excel( ).
  lo_data->create_xml( ).
  lo_data->create_workbook( ).
  lo_data->create_styles( ).
  lo_data->create_worksheet( ).
  lo_data->create_table( ).
*lo_data->create_header_row( ).
  lo_data->add_data_rows( ).
  lo_data->render_xml( ).
  lo_data->download_excel( ).

CLASS lcl_send_mail DEFINITION.

  PUBLIC SECTION.
    INTERFACES:
      yif_prepare_mail.


ENDCLASS.

CLASS lcl_send_mail IMPLEMENTATION.

  METHOD yif_prepare_mail~get_subject.

    rv_subject = TEXT-s00.

  ENDMETHOD.

  METHOD yif_prepare_mail~get_body.

    rt_body = VALUE #( ( line = TEXT-b01 )
                       ( line = cl_abap_char_utilities=>newline )
                       ( line = TEXT-b02 )
                       ( line = cl_abap_char_utilities=>newline )
                       ( line = cl_abap_char_utilities=>newline )
                       ( line = TEXT-b99 && ` ` && sy-sysid && sy-mandt ) ).

  ENDMETHOD.

  METHOD yif_prepare_mail~get_distribution_list.

    TRY.
        DATA(lo_recipient) = cl_distributionlist_bcs=>getu_persistent( i_dliname = iv_dliname
                                                                       i_private = space ).
      CATCH cx_address_bcs INTO DATA(lo_exception).

        DATA(lv_message) = lo_exception->get_longtext( ).

        IF lv_message IS INITIAL.
          lv_message = lo_exception->get_text( ).
        ENDIF.

        MESSAGE lv_message TYPE 'E'.

    ENDTRY.

    APPEND lo_recipient TO rt_recipient.

  ENDMETHOD.

  METHOD yif_prepare_mail~get_attach_file.

    DATA:
      lt_xml  TYPE solix_tab,
      lv_size TYPE i.

    y_send_mail=>tab_to_excel_xml(
      EXPORTING
        ir_data = ir_data   " Any interna table
      CHANGING
        cv_size = lv_size   " Size excel xml
        ct_xml  = lt_xml ). " Excel xml file

* Build file attributes
    rt_file = VALUE #( ( attachment_type     = 'BIN'
                         attachment_subject  = 'Report_Down_payment_' &&
                                               sy-datum+6(2)          &&
                                               sy-datum+4(2)          &&
                                               sy-datum(4)            &&
                                               '_'                    &&
                                               sy-uzeit(4)            &&
                                               '.xlsx'
                         t_file              = lt_xml
                         attachment_size     = lv_size ) ).

  ENDMETHOD.

ENDCLASS.

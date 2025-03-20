*&---------------------------------------------------------------------*
*&      Form  SEND_ERRORS
*&---------------------------------------------------------------------*
FORM send_errors
             USING
               it_log TYPE gtyp_tt_log.

  SELECT grp,
         bname
  FROM yfi_t_vnrq_mail
  INTO TABLE @DATA(lt_vnrq_mail).

  SORT lt_vnrq_mail BY grp ASCENDING.

  IF sy-subrc <> 0.
    RETURN.
  ENDIF.

  DATA:
    lv_message_subject TYPE so_obj_des,
    lt_message_body	   TYPE soli_tab,
    lt_recipient       TYPE rmps_recipient_bcs,
    lt_file	           TYPE yyt_file,
    lv_ref             TYPE REF TO data,
    lt_xml             TYPE solix_tab,
    lv_size            TYPE i,
    lt_ret             TYPE TABLE OF bapiret2,
    lt_addsmtp         TYPE TABLE OF bapiadsmtp,
    lt_log             TYPE gtyp_tt_log.

  lv_message_subject = TEXT-s01.

  lt_message_body = VALUE #( ( line = TEXT-b01 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b02 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b03 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b04 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b05 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b06 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b07 )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = cl_abap_char_utilities=>newline )
                             ( line = TEXT-b99 && ` ` && sy-sysid && sy-mandt ) ).

  LOOP AT lt_vnrq_mail ASSIGNING FIELD-SYMBOL(<fs_vnrq_mail>).

    AT NEW grp.

      CLEAR lt_log.

      lt_log = VALUE #( FOR wa IN it_log WHERE ( grp = <fs_vnrq_mail>-grp ) ( wa ) ).

    ENDAT.

    CHECK lt_log[] IS NOT INITIAL.

    CLEAR:
      lt_recipient,
      lt_file,
      lv_ref,
      lt_xml,
      lv_size,
      lt_ret,
      lt_addsmtp.

    CALL FUNCTION 'BAPI_USER_GET_DETAIL'
      EXPORTING
        username = <fs_vnrq_mail>-bname   " User Name
      TABLES
        return   = lt_ret      " Return Structure
        addsmtp  = lt_addsmtp. " E-Mail Addresses BAPI Structure


    CHECK lt_addsmtp[] IS NOT INITIAL.

    DATA(lv_email) = lt_addsmtp[ 1 ]-e_mail.

* Send email
    TRY.

        DATA(lo_recipient) =
        cl_cam_address_bcs=>create_internet_address( lv_email ).
        APPEND lo_recipient TO lt_recipient.

      CATCH cx_address_bcs.

    ENDTRY.


* Attach file
    lv_ref = REF #( lt_log ).

    y_send_mail=>tab_to_excel_xml(
      EXPORTING
        ir_data = lv_ref    " Any interna table
      CHANGING
        cv_size = lv_size   " Size excel xml
        ct_xml  = lt_xml ). " Excel xml file

* Build file attributes
    lt_file = VALUE #( ( attachment_type    = 'BIN'
                         attachment_subject = 'Vendor_Request_Error_Log_' &&
                                              sy-datum+6(2)               &&
                                              sy-datum+4(2)               &&
                                              sy-datum(4)                 &&
                                              '_'                         &&
                                              sy-uzeit(4)                 &&
                                              '.xlsx'
                         t_file             = lt_xml
                         attachment_size    = lv_size ) ).

    CALL METHOD y_send_mail=>send_mail
      EXPORTING
        iv_sender_user     = sy-uname
        iv_message_subject = lv_message_subject
        it_message_body    = lt_message_body
        it_recipient       = lt_recipient
        it_file            = lt_file.

  ENDLOOP.

ENDFORM.

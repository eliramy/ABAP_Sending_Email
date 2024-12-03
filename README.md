# ABAP_Sending_Email Version 1.0 

From SAP Netweaver 7.0 & above (Include S/4 HANA) 

1) Main method: send_mail .

2) Body can format in text/plain or HTML .

3) Can define specific email or distrubation email also possible control the sender by SAP user name or any fictive email.

4) Can also send the email in future date.

5) Can attach multiples files that can be:

A) binary files converted from spools.

B) AnyForms sapscript/smartforms/adobe converted to binary files.

C) Any files.

Instruction to implement code:
1) file Methods_Desc+Table_Types+Structure_Types.xlsx -> check sheets in file to implement structures & Table types.
2) file y_send_mail.txt -> implement the class .
3) files Interface YIF_PREPARE_MAIL.txt & Example local class send maill attach internal table as excel using interface yif_prepare_mail.txt 
    -> optional example to prepare details for mail interface.
4) all others files example codes to send attach files . 


Enjoy :) 


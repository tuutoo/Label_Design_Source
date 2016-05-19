DOCUMENT "OptioANZ"

  SET TRACE FILE "~errors\OptioANZ.log"
  SET TRACE MASK TRACEMASK("None")
  //SET LOG FILE "~errors\OptioANZ_log.log"

  SET LOG FILE "~errors\OAL_Print.log"
  LET strJobDate = SHOWDATE("%c",NOW()+(3600*8))

  // Variable used to set the location to link with ini file
  LET gLocation = "ANZ"

  // Variables for lookup file
  LET gLaser_OptioPrinter = ""
  LET gLaser_PrinterUNC = ""
  LET gLabel_Folder = ""
  LET gLabel_MFGPro_Output = ""
  LET gLabel_Label_Printer = ""
  LET gLabel_Manufact = ""
  LET gEmailServer = ""
  LET gFaxServer = ""
  LET gEmailFaxOutQ = ""

  // Variables for Company Address lookup
  LET gCompany_Location = ""
  LET gCompany_Name = ""
  LET gCompany_RegNumber = ""
  LET gCompany_Add_Line1 = ""
  LET gCompany_Add_Line2 = ""
  LET gCompany_Add_Line3 = ""
  LET gCompany_Add_Line4 = ""
  LET gCompany_Add_Line5 = ""
  LET gCompany_Add_Line6 = ""
  LET gCompany_Phone = ""
  LET gCompany_Fax = ""
  LET gCompany_Email = ""
  LET gCompany_Web = ""

  //Extract temp file name
  LET ei_tmp = PARAMETER("*",1)
  LET ei_pos = WHERE("~ei",ei_tmp[1])
  LET ei_tmp = ei_tmp[1,(ei_pos+1)->(COLUMNS(ei_tmp[1])-4)]
  
  // Read company address details
  CALL "Read_Company_Details_INI"
  LET gCompany_Location = LookupWS::C_Location
  LET gCompany_Name = LookupWS::C_Name
  LET gCompany_RegNumber = LookupWS::C_RegNumber
  LET gCompany_Add_Line1 = LookupWS::C_Add_Line1
  LET gCompany_Add_Line2 = LookupWS::C_Add_Line2
  LET gCompany_Add_Line3 = LookupWS::C_Add_Line3
  LET gCompany_Add_Line4 = LookupWS::C_Add_Line4
  LET gCompany_Add_Line5 = LookupWS::C_Add_Line5
  LET gCompany_Add_Line6 = LookupWS::C_Add_Line6
  LET gCompany_Phone = LookupWS::C_Phone
  LET gCompany_Fax = LookupWS::C_Fax
  LET gCompany_Email = LookupWS::C_Email
  LET gCompany_Web = LookupWS::C_Web
  
  LET strCompany_Name = LOOKUP(gLocation,gCompany_Location,gCompany_Name)
  LET strCompany_RegNumber = LOOKUP(gLocation,gCompany_Location,gCompany_RegNumber)
  LET strCompany_Add_Line1 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line1)
  LET strCompany_Add_Line2 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line2)
  LET strCompany_Add_Line3 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line3)
  LET strCompany_Add_Line4 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line4)
  LET strCompany_Add_Line5 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line5)
  LET strCompany_Add_Line6 = LOOKUP(gLocation,gCompany_Location,gCompany_Add_Line6)
  LET strCompany_Phone = LOOKUP(gLocation,gCompany_Location,gCompany_Phone)
  LET strCompany_Fax = LOOKUP(gLocation,gCompany_Location,gCompany_Fax)
  LET strCompany_Email = LOOKUP(gLocation,gCompany_Location,gCompany_Email)
  LET strCompany_Web = LOOKUP(gLocation,gCompany_Location,gCompany_Web)

  LET strCompany_Address_Line1 = strCompany_Add_Line1
  LET strCompany_Address_Line2 = strCompany_Add_Line2 & ", " & strCompany_Add_Line3 & ", " & strCompany_Add_Line4
  LET strCompany_Phone_Fax = strCompany_Phone & " " & strCompany_Fax

  // Read printer lookups file to determine print location
  CALL "Read_Printers_INI"
  LET gLaser_OptioPrinter = TOUPPER(LookupWS::OptioPrinter)
  LET gLaser_PrinterUNC = LookupWS::PrinterUNC

  CALL "Read_Label_Printers_INI"
  LET gLabel_Folder = TOUPPER(LookupWS::Folder)
  LET gLabel_MFGPro_Output = TOUPPER(LookupWS::MFGPro_Output)
  LET gLabel_Label_Printer = TOUPPER(LookupWS::Label_Printer)
  LET gLabel_Manufact = TOUPPER(LookupWS::Manufact)

  
  // Read Email and Fax server lookups file to determine server name
  CALL "Read_EmailFaxServer_INI"
  LET gEmailFaxOutQ = LookupWS::EmailFaxQ
  LET gEmailServer = LookupWS::EmailServer
  LET gFaxServer = LookupWS::FaxServer
  
  // Read page 1 for smart select
  SET INPUT "stdin"
  RUN "~doc\AD_GetOutput" RETURNING vOutput_Printer,Is_Hold_Order
  //LOG ("vOutput_Printer: " & vOutput_Printer)
  
  REWIND INPUT

  LET strPrinterUNC = LOOKUP(vOutput_Printer, gLaser_OptioPrinter, gLaser_PrinterUNC)
  LET strLabel_Folder = LOOKUP(vOutput_Printer, gLabel_MFGPro_Output, gLabel_Folder)
  LET strLabel_Label_Printer = LOOKUP(vOutput_Printer, gLabel_MFGPro_Output, gLabel_Label_Printer)
  LET strLabel_Manufact = LOOKUP(vOutput_Printer, gLabel_MFGPro_Output, gLabel_Manufact)
  LET strEmailServer = LOOKUP(voutput_Printer, gEmailFaxOutQ, gEmailServer)
  LET strFaxServer = LOOKUP(voutput_Printer, gEmailFaxOutQ, gFaxServer)

  //LOG ("strPrinterUNC: " & strPrinterUNC)

  READ INPUT

  // Fields to use for detecting form type
  LET Is_DeliveryDocket = PRUNE(@[2->3,1->7])
  LET Is_OrderAcknowledgment = SQUISH(PRUNE(@[6->12,1->6]))
  LET Is_Statement = WHERE("CUSTOMER STATEMENT", TOUPPER(@[1]))
  LET Is_Remittance = WHERE("REMITTANCE ADVICE", TOUPPER(@[1]))
  LET Is_ClaimsInvoice = WHERE("CLAIMS INVOICE", TOUPPER(@[1]))

  LET Is_CustomsInvoice = TOUPPER(PRUNE(@[9->12,20->28]))
  LET Is_CustomsInvoice1 = WHERE("CTRY",TOUPPER(@))
  LET Is_Invoice = PRUNE(@[5,44->110])
  LET Is_con = PRUNE(@[5,44->110])
  LET Is_ExportInvoice = PRUNE(@[5,44->110])
  LET Is_Invoice_ANZ = PRUNE(@[26->40,67->81])
  LET Is_Stocktake_Label = WHERE("TAG NUMBER",TOUPPER(@[1]))
  LET Is_PickList = WHERE("PICKLIST PRINT", TOUPPER(PRUNE(@[1])))
  LET Is_CreditMemo = WHERE("MEMO", TOUPPER(@[7]))
  LET Is_COCDoc = WHERE("COC DOCUMENT", TOUPPER(@[4]))  
  LET Is_COCDoc2 = WHERE("COC DOCUMENT", TOUPPER(@[3]))  
  LET Is_ADPallet = WHERE("COMPANY LOGO", TOUPPER(PRUNE(@[2]))) // Jason Added 
  LET Is_PalletType2 =  WHERE("NO OF LOTS", TOUPPER(PRUNE(@[14]))) // Jason Added 

  // Loop through all rows in Is_Invoice_ANZ
  FOR iLoopCounter = 1 TO ROWS(Is_Invoice_ANZ)
    IF WHERE("Extended Price", Is_Invoice_ANZ[iLoopCounter]) THEN
      LET b_Is_Invoice_ANZ = 1
    END IF
  END FOR

  //Labels
  LET Is_InternalLabel_Extract = PRUNE(@[2,15->30])
  LET Is_InternalLabel = PRUNE(@[1])
  LET gquote = "\""
  LET Is_MultiLabel = PRUNE(@[1,1->5])

  // Field for purchase order
  	LET Is_PurchaseOrder = PRUNE(@[5,53->69])
  IF Is_PurchaseOrder = "" THEN
  	READ INPUT
  	LET Is_PurchaseOrder = PRUNE(@[5,53->69]) 
  END IF

  // Rewind spool data ready for corresponding form
  REWIND INPUT

  IF  Is_con = "CONSOLIDATED" THEN
    //LOG "Invoice printed"
    RUN "~doc\ANZ\AD_TaxInvoicemod" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Consolidated" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)

ELIF (Is_CustomsInvoice[1] = "CTRY OF" OR Is_CustomsInvoice[2] = "CTRY OF" OR Is_CustomsInvoice[3] = "CTRY OF" OR Is_CustomsInvoice[4] = "CTRY OF") THEN
    RUN "~doc\ANZ\AD_CustomsInvoice" WITH strPrinterUNC, "FALSE", ei_tmp 

ELIF Is_CreditMemo > 0 THEN
    RUN "~doc\ANZ\AD_TaxAdjustment" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Credit Memo" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
ELIF Is_COCDoc > 0 OR Is_COCDoc2 > 0 THEN
    RUN "~doc\ANZ\AD_COCDoc" WITH strPrinterUNC 
ELIF Is_Invoice = "I N V O I C E" THEN
    //LOG "Invoice printed"
    RUN "~doc\ANZ\AD_TaxInvoice" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Tax Invoice" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_ExportInvoice = "EXPORT INVOICE" THEN
    //LOG "Export Invoice printed"
    RUN "~doc\ANZ\AD_ExportInvoice" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Export Invoice" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_ClaimsInvoice > 0 THEN
    RUN "~doc\ANZ\AD_ClaimsInvoice" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Claims Invoice" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)

// Jason Added 
// ELIF Is_ADPallet > 0 THEN    
//          RUN "~doc\ANZ\AD_PalletDoc"   WITH strPrinterUNC  , "FALSE",  ei_tmp

  ELIF Is_Statement > 0 THEN
    RUN "~doc\ANZ\AD_Statement" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Statement" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_Remittance > 0 THEN
    RUN "~doc\ANZ\AD_Remittance" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Remittance" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_PickList > 0 THEN
     RUN "~doc\ANZ\AD_PickList" WITH strPrinterUNC,strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Pick List" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
   // Lee Li added 
  ELIF  Is_Hold_Order > 0 THEN
    RUN "~doc\ANZ\AD_OrderAck_HD" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
   // Lee Li added
  ELIF Where_Array("CHANGE", TOUPPER(Is_OrderAcknowledgment)) = "TRUE"  THEN
    RUN "~doc\ANZ\AD_OrderAck" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Order Ack" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_PurchaseOrder = "P U R C H A S E" THEN
    RUN "~doc\ANZ\AD_PurchaseOrder" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Purchase Order" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELIF Is_MultiLabel = gquote & "Slit" OR Is_MultiLabel = gquote & "Shee" OR Is_MultiLabel = gquote & "Grap" OR Is_MultiLabel = gquote & "Inte" THEN 
    RUN "~doc\ANZ\MultipleLabels" WITH strLabel_Folder, strLabel_Label_Printer, strLabel_Manufact
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    IF Is_MultiLabel = gquote & "Slit" THEN
      LET strJobType = "Slit"
    ELIF Is_MultiLabel = gquote & "Shee" THEN 
      LET strJobType = "Sheet"
    ELIF Is_MultiLabel = gquote & "Grap" THEN 
      LET strJobType = "Graphic"
    ELIF Is_MultiLabel = gquote & "Inte" THEN 
      LET strJobType = "Internal"
    END IF

    LOG (strJobType & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strLabel_Label_Printer)
  //jerryg add 130707
  ELIF Is_PalletType2 > 0 THEN
    RUN "~doc\ANZ\PalletLabelsType2" WITH strLabel_Folder 

  ELIF Is_Stocktake_Label > 0 THEN
    RUN "~doc\ANZ\Stocktake" WITH strPrinterUNC, ei_tmp
 
//  ELIF b_Is_Invoice_ANZ THEN
//	RUN "~doc\ANZ\AD_TaxInvoice_ANZ" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
  ELIF (Is_DeliveryDocket[1] = "BILL-TO" OR Is_DeliveryDocket[2] = "BILL-TO") THEN
    RUN "~doc\ANZ\AD_DeliveryDocket" WITH strPrinterUNC, "FALSE", ei_tmp, strCompany_Name, strCompany_Address_Line1, strCompany_Address_Line2, strCompany_Phone_Fax
    LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
    LOG ("Delivery Docket" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strPrinterUNC)
  ELSE 	
    RUN "~doc\ANZ\PalletLabels"

   //jerryg mark start 130707
   //RUN "~doc\ANZ\AD_PalletDoc"   WITH strPrinterUNC  , "FALSE",  ei_tmp
  //  LET strJobEndDate = SHOWDATE("%c",NOW()+(3600*8))
  //  IF strPrinterUNC <> "" AND strPrinterUNC <> "OPTIO_E" THEN
  //    LOG ("Pallet" & "," & strJobDate & "," & strJobEndDate & "," & ei_tmp & "," & strLabel_Label_Printer)
  //  END IF
  //jerryg mark end 130707
  END IF  
  
END DOCUMENT

CHANNEL "Email Fax Server"
  FILE "~config\ANZ_AD_EmailFaxServer.ini" MODE "READ"
  ROWS 1
  COLUMNS 0
  RETURNS "YES"
  LINEFEEDS "YES"
  TABS "NO"
  FIELDS EmailFaxQ, EmailServer, FaxServer
  SEPARATOR COMMAS
  DELIMITER QUOTES
END CHANNEL

CHANNEL "Printers"
  FILE "~config\ANZ_AD_Printers.ini" MODE "READ"
  ROWS 1
  COLUMNS 0
  RETURNS "YES"
  LINEFEEDS "YES"
  TABS "NO"
  FIELDS OptioPrinter, PrinterUNC
  SEPARATOR COMMAS
  DELIMITER QUOTES
END CHANNEL

CHANNEL "Label Printers"
  FILE "~config\ANZ_AveryDennisonPrinters.ini" MODE "READ"
  ROWS 1
  COLUMNS 0
  RETURNS "YES"
  LINEFEEDS "YES"
  TABS "NO"
  FIELDS Folder, MFGPro_Output, Label_Printer, Printer_IP, Make_Model, Manufact
  SEPARATOR COMMAS
  DELIMITER QUOTES
END CHANNEL

CHANNEL "Company Address Details"
  FILE "~config\Avery_Company_Details.ini" MODE "READ"
  ROWS 1
  COLUMNS 0
  RETURNS "YES"
  LINEFEEDS "YES"
  TABS "NO"
  FIELDS C_Location, C_Name, C_RegNumber, C_Add_Line1, C_Add_Line2, C_Add_Line3, C_Add_Line4, C_Add_Line5, C_Add_Line6, C_Phone, C_Fax, C_Email, C_Web
  SEPARATOR COMMAS
  DELIMITER QUOTES
END CHANNEL

FUNCTION "Read_Company_Details_INI"
  DESTROY WORKSHEET "LookupWS"
  CREATE WORKSHEET "LookupWS"
  OPEN "Company Address Details"
  WHILE NOT EOF("Company Address Details")
    READ "Company Address Details"
    LET LookupWS::C_Location[0] = C_Location
    LET LookupWS::C_Name[0] = C_Name
    LET LookupWS::C_RegNumber[0] = C_RegNumber
    LET LookupWS::C_Add_Line1[0] = C_Add_Line1
    LET LookupWS::C_Add_Line2[0] = C_Add_Line2
    LET LookupWS::C_Add_Line3[0] = C_Add_Line3
    LET LookupWS::C_Add_Line4[0] = C_Add_Line4
    LET LookupWS::C_Add_Line5[0] = C_Add_Line5
    LET LookupWS::C_Add_Line6[0] = C_Add_Line6
    LET LookupWS::C_Phone[0] = C_Phone
    LET LookupWS::C_Fax[0] = C_Fax
    LET LookupWS::C_Email[0] = C_Email
    LET LookupWS::C_Web[0] = C_Web
  END WHILE
  CLOSE "Company Address Details"
END FUNCTION

FUNCTION "Read_Printers_INI"
  DESTROY WORKSHEET "LookupWS"
  CREATE WORKSHEET "LookupWS"
  OPEN "Printers"
  WHILE NOT EOF("Printers")
    READ "Printers"
    LET LookupWS::OptioPrinter[0] = OptioPrinter
    LET LookupWS::PrinterUNC[0] = PrinterUNC
  END WHILE
  CLOSE "Printers"
END FUNCTION

FUNCTION "Read_Label_Printers_INI"
  DESTROY WORKSHEET "LookupWS"
  CREATE WORKSHEET "LookupWS"
  OPEN "Label Printers"
  WHILE NOT EOF("Label Printers")
    READ "Label Printers"
    LET LookupWS::Folder[0] = Folder
    LET LookupWS::MFGPro_Output[0] = MFGPro_Output
    LET LookupWS::Label_Printer[0] = Label_Printer
    LET LookupWS::Manufact[0] = Manufact
  END WHILE
  CLOSE "Printers"
END FUNCTION

FUNCTION "Read_EmailFaxServer_INI"
  DESTROY WORKSHEET "LookupWS"
  CREATE WORKSHEET "LookupWS"
  OPEN "Email Fax Server"
  WHILE NOT EOF("Email Fax Server")
    READ "Email Fax Server"
    LET LookupWS::EmailFaxQ[0] = EmailFaxQ
    LET LookupWS::EmailServer[0] = EmailServer
    LET LookupWS::FaxServer[0] = FaxServer
  END WHILE
  CLOSE "Email Fax Server"
END FUNCTION

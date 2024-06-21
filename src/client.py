import imaplib
import email
import os
import json
import csv
import pytesseract
import pdf2image

import gemini

from enum import Enum
from pytesseract import Output
from email.header import decode_header

class FormType(Enum):
    SALE_QUOTATION = 1
    AUTHORIZATION = 2

def pdfToImage(pdf_bytes: bytes, poppler_path: str = 'poppler/bin'):
    images = pdf2image.convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
    
    result = []
    
    for img_data in enumerate(images):
        result.append(rotateWithOSD(img_data[1]))
    
    return result

def rotateWithOSD(image):
    text = pytesseract.image_to_string(image)
    
    if text != '':
        osd = pytesseract.image_to_osd(image, output_type=Output.DICT, config='--psm 0')
        
        if osd['orientation'] != 0:
            # rotates image back to normal if the image is rotated in the pdf
            image = image.rotate(osd['orientation'], expand=True)
    
    return image

def recursiveListUpdate(current_data: list, new_data: list = []):
    for data in current_data:
        if isinstance(data, list):
            recursiveListUpdate(data, new_data)
        else:
            new_data.append(data)
    
    return new_data

def setTesseractPath(path):
    pytesseract.pytesseract.tesseract_cmd = path

class EmailClient:
    def __decodeH(self, header):
        header_part, encoding = decode_header(header)[0]
        if isinstance(header_part, bytes):
            header_part = header_part.decode(encoding, errors='replace') # if it's a bytes, decode to str
        return header_part

    def __fetch_email(self, index: int, message_parts: str = "(RFC822)"):
        res, msg = self.__imap.fetch(str(index), message_parts)
        
        responses = []
        
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                mail = email.message_from_bytes(response[1])
                
                # decode the email subject and sender
                subject = self.__decodeH(mail.get("Subject"))
                sender_address = self.__decodeH(mail.get("Return-Path"))
                
                body = ""
                
                attachments = []
                    
                # if the email message is multipart
                if mail.is_multipart():
                    # iterate over email parts
                    
                    part_index = 0
                    
                    for part in mail.walk():
                        # skip the first part (the email itself)
                        if part_index == 0:
                            part_index += 1
                            continue
                        
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        
                        # download attachment
                        if content_type == "application/pdf" and "attachment" in content_disposition:
                            data = part.get_payload(decode=True)
                            
                            if data != None:
                                attachments.append(pdfToImage(data))
                                
                        elif (content_type == "image/jpeg" or content_type == "image/png") and "attachment" in content_disposition:
                            data = part.get_payload(decode=True)
                            
                            if data != None:
                                attachments.append(rotateWithOSD(data))
                        
                        elif content_type == "text/plain":
                            body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                    
                        part_index += 1
                        
                else:
                    # extract the email body
                    body = mail.get_payload(decode=True).decode('utf-8', errors='replace')
                
                attachments = recursiveListUpdate(attachments)
                
                return {
                    "subject": subject,
                    "from": sender_address,
                    "body": body,
                    "attachments": attachments
                }

    def fetchAndParse(self, save_path_folder: str = "outputs"):
        # update the number of emails in the inbox
        status, messages = self.__imap.search(None, "ALL")
        self.total_email_count = len(messages[0].split())
        
        # calculate the number of new emails
        self.new_email_count = self.total_email_count - self.current_email
        
        # fetch the new emails
        for i in range(self.new_email_count):
            index = self.total_email_count - i
            
            print("Fetching email-" + str(index) + "...") # log the fetching of the email
            currentMail = self.__fetch_email(index)
            
            if currentMail["body"].strip() == "":
                print("Email-" + str(index) + " has no body. Skipping...")
                continue
            
            si_s1 = r'You are an expert of reading emails. Your job is extracting item/part information in json format. Completely ignore the rest of the email and do not add "´" characters. Example result: {"Vendor name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}'
            si_s2_filtering = r'You are an expert of reading documents. Your job is giving type every document that we are going to give you. We have 3 types of documents: "Sale Quotation Form-1", "Authorization Form-2" and "Other-3". You need to detect the type of the document and return it. Do not add "´" characters to the result. Example result: {"Type": "Sale Quotation Form-1"}. Here is the tricky part: these documents can contain unnecessary information as well so you need to be aware of which information is related or not.'
            si_s3_sale = r'You have a wide range of information about sale quotation forms. Your job is extracting item/part information from the forms that we are going to give you. Do not add "´" characters to the result. Example result: {"Vendor name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}. DO NOT forget that these keywords are just examples so you need to extract every information that is related to the part/item. Here is the tricky part: these forms can contain unnecessary information as well so you need to be aware of which information is related or not.'
            si_s3_auth = r'You have a wide range of information about authorization forms. Your job is extracting necessary information about the form that we are going to give you. Do not add "´" characters to the result. Example result: {"Tagged by": "${company_name}", "Tag date": "1234", "Tag type(FAA/EASA)": "1234", "Dual": "1234"}. To detect whether the authorization form dual or single you need to check if "Other regulation specified in Block 12" is marked or not. If it is marked, the form is dual.'
            si_s4 = r'Your job is merging json files that are given to you. You need to merge them in a way that the result is a single json file. Do not add "´" characters to the result. Here is the tricky part: Some of the keys or values may be same in different files. You need to be aware of this and merge the same elements in the files. Example result: {"Vendor_name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}'
            si_s5 = r'Normalize the json data. Remove null values and unnecessary keys. Then give the result without adding "`" or "json".'
            
            # Step 1 - Extracting item/part information in json format
            print("Step - 1")
            gemini.setSystemInstructions(si_s1)
            mail_info = gemini.generateContent(currentMail["body"])
            mail_info = mail_info.replace("`", "")
            mail_info = mail_info.replace("json", "")
            mail_info = json.loads(mail_info)

            # Step 2 - Detecting the type of the document
            print("Step - 2")
            gemini.setSystemInstructions(si_s2_filtering)
            
            attachmentsInfo = []
            
            for a in currentMail["attachments"]:
                attachment_type = gemini.generateContent("document: ", [a])
                
                if attachment_type.find("1") != -1:
                    form_type = FormType.SALE_QUOTATION
                elif attachment_type.find("2") != -1:
                    form_type = FormType.AUTHORIZATION
                else:
                    continue
                    
                attachmentsInfo.append({"attachment": a, "type": form_type, "info": {}})
                    
            # Step 3 - Extracting item/part information from the forms
            print("Step - 3")
            for aInfo in attachmentsInfo:
                if aInfo["type"] == FormType.SALE_QUOTATION:
                    gemini.setSystemInstructions(si_s3_sale)
                    info = gemini.generateContent("document: ", [aInfo["attachment"]])
                elif aInfo["type"] == FormType.AUTHORIZATION:
                    gemini.setSystemInstructions(si_s3_auth)
                    info =  gemini.generateContent("document: ", [aInfo["attachment"]])
                else:
                    continue
                
                print("\n" + info + "\n")
                
                # Remove '`' and 'json' from the result
                info = info.replace("`", "")
                info = info.replace("json", "")
                
                aInfo["info"] = json.loads(info)
            
            # Step 4 - Merging json files
            print("Step - 4")
            gemini.setSystemInstructions(si_s4)
            
            prompt = "Json data to merge:" + json.dumps(mail_info) + ","
            
            for attachment in attachmentsInfo:
                prompt += json.dumps(attachment["info"]) + ","
                
            merged_info = gemini.generateContent(prompt)
            
            print("\n" + merged_info + "\n")

            # Step 5 - Normalize the json data
            print("Step - 5")
            
            # remove the unnecessary keys
            required_keys = "vendor_name, part_no, cond(CC or Cond), qty(QTY or Quantity), lead_time, price, description, serial_number, notes, warranty, dual, tagged_by, tag_date, tag_type(FAA/EASA), stock_type(EA, OH, SV, RP)"
            
            gemini.setSystemInstructions(si_s5)
            
            prompt = "Required keys: " + required_keys + ", Json data: " + json.dumps(merged_info)
            
            merged_info = gemini.generateContent(prompt)
            
            # Remove '`' and 'json' from the result
            merged_info = merged_info.replace("`", "")
            merged_info = merged_info.replace("json", "")
            
            merged_info = json.loads(merged_info)
            
            # Save as CSV
            if not os.path.isdir(save_path_folder):
                os.makedirs(save_path_folder)
                
            file_path = os.path.join(save_path_folder, "email-" + str(index) + ".csv")
            
            print("Saving into: " + file_path)
            
            self.saveAsCSV(merged_info, file_path)
        
    def selectMailbox(self, mailbox: str):
        print("Selecting mailbox " + mailbox + "...")
        self.__imap.select(mailbox)
        
    def login(self, username: str, password: str):
        print("Logging in to email...")
        self.__imap.login(username, password)
        
    def setCurrentEmailIndex(self, index: int):
        print("Setting current email index to " + str(index) + "...")
        self.current_email = index
        
    def connectIMAP(self, imap_server: str, imap_ssl_port: int = 993):
        print("Connecting to IMAP server...")
        self.__imap = imaplib.IMAP4_SSL(imap_server, imap_ssl_port)

    def logoutAndClose(self):
        print("Logging out and closing email client...")
        self.__imap.close()
        self.__imap.logout()
        
    def saveAsCSV(self, data: dict, save_path: str):
        with open(os.path.normpath(save_path), 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(data.keys())  # write header
            writer.writerow(data.values())  # write rows

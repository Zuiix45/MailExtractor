import imaplib
import email
import os
import json
import csv
import re
import requests
import pytesseract
import pdf2image

import gemini

from time import sleep
from enum import Enum
from pytesseract import Output
from email.header import decode_header

class FormType(Enum):
    SALE_QUOTATION = 1
    AUTHORIZATION = 2

def saveImageObject(image_object, save_path_folder: str = "outputs"):
    if not os.path.isdir(save_path_folder):
        os.makedirs(save_path_folder)
        
    file_path = os.path.join(save_path_folder, "email-" + str(image_object["email"]) + "-attachment-" + str(image_object["attachment"]) + "-page-" + str(image_object["page"]) + ".png")
    
    image_object["image_data"].save(file_path, "PNG")
    
    return file_path

def saveAsCSV(data: dict, save_name: str, save_path_folder: str = "outputs"):
    try:
        with open(os.path.normpath(os.path.join(save_path_folder, save_name)), 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(data.keys())  # write header
            writer.writerow(data.values())  # write rows
    except Exception as e:
        print(f"Error saving data to CSV\n data: {data}\n error: {e}")

def pdfToImage(pdf_bytes: bytes, poppler_path: str = 'poppler/bin'):
    try:
        images = pdf2image.convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
    except pdf2image.exceptions.PDFPageCountError:
        return []
    
    result = []
    
    for img_data in enumerate(images):
        result.append(img_data[1])
    
    return result

def getOCD(image_object):
    try:
        image_object["text"] = pytesseract.image_to_string(image_object["image_data"])
        
        if image_object["text"] != '':
            image_object["osd"] = pytesseract.image_to_osd(image_object["image_data"], output_type=Output.DICT, config='--psm 0')
            
    except pytesseract.pytesseract.TesseractError:
        print(f"Image: e{image_object['email']}-a{image_object['attachment']}-p{image_object['page']} could not be processed.")
    
    return image_object

def rotateImage(image_object):
    if "orientation" in image_object['osd']: 
        if image_object['osd']['orientation'] != 0:
            # rotates image back to normal if the image is rotated in the pdf
            image_object['image_data'] = image_object['image_data'].rotate(image_object['osd']['orientation'], expand=True)
    else:
        print(f"Image: e{image_object['email']}-a{image_object['attachment']}-p{image_object['page']} could not be rotated.")
    
    return image_object

def recursiveListUpdate(current_data: list, new_data: list = []):
    for data in current_data:
        if isinstance(data, list):
            recursiveListUpdate(data, new_data)
        else:
            new_data.append(data)
    
    return new_data

def setTesseractPath(path):
    pytesseract.pytesseract.tesseract_cmd = path

def step1(mail_body, system_instructions: str):
    gemini.setSystemInstructions(system_instructions)
    mail_info = gemini.generateContent(mail_body) # request count: 1
    mail_info = mail_info.replace("`", "")
    mail_info = mail_info.replace("json", "")

    parts = mail_info.split("\n")

    parts_data = []

    for part in parts:
        if part.strip():  # Skip empty parts
            try:
                parts_data.append(json.loads(part))
            except Exception as e:
                print(f"Error parsing part-{parts.index(part)+1} - Error:", e)

    return parts_data

def step2(attachments, system_instructions: str):
    gemini.setSystemInstructions(system_instructions)
    
    attachmentsInfo = []
    
    for attachment in attachments:
        attachment_type = gemini.generateContent("document: ", [attachment["image_data"]]) # request count: a
        
        if attachment_type.find("1") != -1:
            form_type = FormType.SALE_QUOTATION
        elif attachment_type.find("2") != -1:
            form_type = FormType.AUTHORIZATION
        else:
            continue
            
        attachmentsInfo.append({"attachment": attachment, "type": form_type, "info": {}})
        
    return attachmentsInfo

def step3(attachmentInfo, sale_quotation_instructions: str, authorization_instructions: str):
    if attachmentInfo["type"] == FormType.SALE_QUOTATION:
        gemini.setSystemInstructions(sale_quotation_instructions)
        info = gemini.generateContent("document: ", [attachmentInfo["attachment"]["image_data"]]) # request count: 1
    elif attachmentInfo["type"] == FormType.AUTHORIZATION:
        gemini.setSystemInstructions(authorization_instructions)
        info =  gemini.generateContent("document: ", [attachmentInfo["attachment"]["image_data"]]) # request count: 1
    
    # Remove '`' and 'json' from the result
    info = info.replace("`", "")
    info = info.replace("json", "")
    
    attachmentInfo["info"] = json.loads(info)
        
    return attachmentInfo

def step4(partInfo, attachmentsInfo, system_instructions: str):
    gemini.setSystemInstructions(system_instructions)
    
    prompt = "Json data to merge:" + json.dumps(partInfo) + ","
    
    for attachment in attachmentsInfo:
        prompt += json.dumps(attachment["info"]) + ","
        
    merged_info = gemini.generateContent(prompt) # request count: 1
    
    return merged_info

def step5(merged_info, mail_index: int, part_index: int, system_instructions: str):
    # remove the unnecessary keys
    required_keys = "vendor_name, part_no, cond(CC or Cond), qty(QTY or Quantity), lead_time, price, description, serial_number, notes, warranty, dual, tagged_by, trace_to, tag_date, tag_type(FAA/EASA), stock_type(EA, OH, SV, RP)"
    
    gemini.setSystemInstructions(system_instructions)
    
    prompt = "Required keys: " + required_keys + ", Json data: " + json.dumps(merged_info)
    
    normalized_info = gemini.generateContent(prompt) # request count: 1
    
    # Remove '`' and 'json' from the result
    normalized_info = normalized_info.replace("`", "")
    normalized_info = normalized_info.replace("json", "")
    
    try:
        normalized_info = json.loads(normalized_info)
    except json.JSONDecodeError:
        print(f"Json data could not be parsed for e{mail_index}-p{part_index+1}. You may check output manually: {normalized_info}")
        normalized_info = {}
    finally:
        return normalized_info

class EmailClient:
    def __init__(self):
        self.isLogged_in = False
    
    def __decodeH(self, header):
        header_part, encoding = decode_header(header)[0]
        if isinstance(header_part, bytes):
            header_part = header_part.decode(encoding, errors='replace') # if it's a bytes, decode to str
        return header_part

    def __parse_image_data(self, index, attachment_index, page_index, image_data):
        image_object = {"email": index, "attachment": attachment_index, "page": page_index+1, "image_data": image_data, "text": "", "osd": {}}
        image_object = getOCD(image_object)
        
        # Check if image contains text
        if image_object["text"].strip() != "":
            return rotateImage(image_object)

    def __fetch_email(self, index: int, message_parts: str = "(RFC822)"):
        status, msg = self.__imap.fetch(str(index), message_parts)
        
        if status != "OK":
            print("An error occurred while fetching the email.")
            return
        
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
                    attachment_index = 1
                    
                    for part in mail.walk():
                        # skip the first part (the email itself)
                        if part_index == 0:
                            part_index += 1
                            continue
                        
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        
                        # parse pdf attachments
                        if content_type == "application/pdf" and "attachment" in content_disposition:
                            data = part.get_payload(decode=True)
                            
                            if data != None:
                                temp_image_data = pdfToImage(data)
                                
                                for di in range(len(temp_image_data)):
                                    image_object = self.__parse_image_data(index, attachment_index, di+1, temp_image_data[di])
                                    
                                    if image_object != None:
                                        attachments.append(image_object)
                                
                            attachment_index += 1
                        
                        # parse image attachments
                        elif (content_type == "image/jpeg" or content_type == "image/png") and "attachment" in content_disposition:
                            data = part.get_payload(decode=True)
                            
                            if data != None:
                                image_object = self.__parse_image_data(index, attachment_index, 0, data)
                                
                                if image_object != None:
                                    attachments.append(image_object)
                                
                            attachment_index += 1
                        
                        elif content_type == "text/plain":
                            body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                    
                        part_index += 1
                        
                else:
                    # extract the email body for single part messages
                    body = mail.get_payload(decode=True).decode('utf-8', errors='replace')
                
                subject_output = subject
                sender_address_output = sender_address
                body_output = body
                attachments_output = recursiveListUpdate(attachments)
                
                pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
                links = re.findall(pattern, body)
                
                # Download the files attached to the link
                for link in links:
                    if link.endswith('>'):
                        link = link[:-1]
                    
                    response = requests.get(link, stream=True)
                    
                    if response.status_code == 200:
                        if response.content != None:
                            images = pdfToImage(response.content)
                            
                            if len(images) > 0:
                                for di in range(len(images)):
                                    image_object = self.__parse_image_data(index, attachment_index, di+1, images[di])
                                    
                                    if image_object != None:
                                        attachments.append(image_object)
                
                print(f"Total {len(attachments)} attachments found.")
                
                # reset values
                subject = ""
                sender_address = ""
                body = ""
                attachments = []
                
                return {
                    "subject": subject_output,
                    "from": sender_address_output,
                    "body": body_output,
                    "attachments": attachments_output
                }

    def getTotalEmailCount(self):
        status, messages = self.__imap.search(None, "ALL")
        self.total_email_count = len(messages[0].split())
        return self.total_email_count

    def fetchAndParse(self, mail_index, save_path_folder: str = "outputs", sleep_duration_after_finish: int = 10):
        print("\nFetching email-" + str(mail_index) + "...") # log the fetching of the email
        currentMail = self.__fetch_email(mail_index)
        
        email_save_folder = os.path.join(save_path_folder, "email-" + str(mail_index))
        
        # Create folder to save outputs
        if not os.path.isdir(email_save_folder):
            os.makedirs(email_save_folder)
        
        ### Save Attachments ###
        for attachment in currentMail["attachments"]:
            saveImageObject(attachment, email_save_folder)
        
        # Starting AI Workflow
        if currentMail["body"].strip() == "":
            print("Email-" + str(mail_index) + " has no body. Skipping...")
            return
        
        si_s1 = r'You are an expert of reading emails. Your job is extracting item/part information in json format. Completely ignore the rest of the email and do not add "´" characters. Example result: {"Vendor name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}. If there is a table of parts... give the result as a list of jsons separated by new lines. But there is a low chance of having a table.'
        si_s2_filtering = r'You are an expert of reading documents. Your job is giving type every document that we are going to give you. We have 3 types of documents: "Sale Quotation Form-1", "Authorization Form-2" and "Other-3". You need to detect the type of the document and return it. Do not add "´" characters to the result. Example result: {"Type": "Sale Quotation Form-1"}. Here is the tricky part: these documents can contain unnecessary information as well so you need to be aware of which information is related or not.'
        si_s3_sale = r'You have a wide range of information about sale quotation forms. Your job is extracting item/part information from the forms that we are going to give you. Do not add "´" characters to the result. Example result: {"Vendor name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}. DO NOT forget that these keywords are just examples so you need to extract every information that is related to the part/item. Here is the tricky part: these forms can contain unnecessary information as well so you need to be aware of which information is related or not.'
        si_s3_auth = r'You have a wide range of information about authorization forms. Your job is extracting necessary information about the form that we are going to give you. Do not add "´" characters to the result. Example result: {"Tagged by": "${company_name}", "Tag date": "1234", "Tag type(FAA/EASA)": "1234", "Dual": "1234"}. To detect whether the authorization form dual or single you need to check if "Other regulation specified in Block 12" is marked or not. If it is marked, the form is dual.'
        si_s4 = r'Your job is merging json files that are given to you. You need to merge them in a way that the result is a single json file. Do not add "´" characters to the result. Here is the tricky part: Some of the keys or values may be same in different files. You need to be aware of this and merge the same elements in the files. Example result: {"Vendor_name": "1234", "Part No": "1234", "CC": "1234", "QTY": "1234", "Lead Time": "1234", "Unit Price (USD)": "1234", "description": "1234", "Serial number": "1234 ", "Trace To": "1234", "Tag type": "1234", "Tagged by": "1234", "Notes": "1234", "Warranty": "1234", "Dual": "1234"}'
        si_s5 = r'Normalize the json data. Remove null values and unnecessary keys. Then give the result without adding "`" or "json".'
        
        # Step 1 - Extracting item/part information in json format - total request count: 1
        print("Step 1 - Extracting item/part information in json format...")
        parts_data = step1(currentMail["body"], si_s1)
        
        if len(parts_data) == 0:
            print("No part information could be extracted from the mail body. Skipping...")
            return
        
        print(f"Total parts extracted from mail body: {len(parts_data)}")

        # Step 2 - Detecting the type of the attachments - total request count: a
        print("Step 2 - Detecting the type of the attachments...")
        attachmentsInfo = step2(currentMail["attachments"], si_s2_filtering)

        # Step 3 - Extracting item/part information from the attachments - total request count: a
        print("Step 3 - Extracting item/part information from the attachments...")
        
        parsed_attachments = []
        
        for aInfo in attachmentsInfo:
            parsed_attachments.append(step3(aInfo, si_s3_sale, si_s3_auth))
        
        # Step 4 - Merging json files - total request count: 1
        print("Step 4 - Merging json files...")
        
        for part in parts_data:
            part_index = parts_data.index(part)
            merged_info = step4(part, parsed_attachments, si_s4)
            
            # Step 5 - Normalize the json data - total request count: 1
            print(f"Step 5 - Normalizing the part-{part_index+1} data...")
            normalized_info = step5(merged_info, mail_index, part_index, si_s5)
            
            if len(normalized_info) == 0:
                print(f"Part-{part_index+1} data could not be normalized. Skipping...")
                continue
            
            print(f"Saving part-{part_index+1} data...")
            saveAsCSV(normalized_info, f"e{mail_index}-p{part_index+1}.csv", email_save_folder)
            
        print("Email-" + str(mail_index) + " has been processed successfully.")
        
        sleep(sleep_duration_after_finish)
        
    def selectMailbox(self, mailbox: str):
        self.__imap.select(mailbox)
        
        print("Selecting mailbox " + mailbox + "...")
        
    def login(self, username: str, password: str):
        self.__imap.login(username, password)
        self.isLogged_in = True
        
        print("Logging in to email client...")
        
    def setCurrentEmail(self, num: int):
        self.current_email = num-1
        
        print("Setting current email index to " + str(num) + "...")
        
    def connectIMAP(self, imap_server: str, imap_ssl_port: int = 993):
        self.__imap = imaplib.IMAP4_SSL(imap_server, imap_ssl_port)
        
        print("Connecting to IMAP server " + imap_server + "...")

    def logoutAndClose(self):
        self.__imap.close()
        self.__imap.logout()
        
        print("Logging out and closing email client...")

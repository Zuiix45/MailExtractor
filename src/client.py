import imaplib
import email
import asyncio
import os
import json
import csv
import google.generativeai as genai

from datetime import datetime
from email.header import decode_header
from google.generativeai.types import HarmCategory, HarmBlockThreshold

class EmailClient:
    def __init__(self, username: str, password: str, imap_server: str, imap_ssl_port: int = 993, starting_email_index: int = 0):
        self.__imap = imaplib.IMAP4_SSL(imap_server, imap_ssl_port)
        self.email_history = []
        self.new_emails = []
        self.total_email_count = starting_email_index
        self.new_email_count = 0
        
        print("Logging in to email server...")
    
        self.__imap.login(username, password)
        self.__imap.select("INBOX")
        
    def __decodeH(self, header):
        header_part, encoding = decode_header(header)[0]
        if isinstance(header_part, bytes):
            header_part = header_part.decode(encoding) # if it's a bytes, decode to str
        return header_part

    def __fetch_email(self, index: int, message_parts: str = "(RFC822)"):
        res, msg = self.__imap.fetch(str(index), message_parts)
        
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                
                # decode the email subject and sender
                subject = self.__decodeH(msg.get("Subject"))
                sender_address = self.__decodeH(msg.get("Return-Path"))
                    
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        
                        # download attachment
                        if "attachment" in content_disposition:
                            filename = part.get_filename()
                            
                            if filename:
                                data = part.get_payload(decode=True)
                                
                                if data != None:
                                    if not os.path.isdir("attachments"):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir("attachments")
                                    
                                    filepath = os.path.join("attachments", filename)
                                    
                                    # download attachment and save it
                                    open(filepath, "wb").write(data)
                        
                        elif content_type == "text/plain":
                            body = part.get_payload(decode=True).decode()
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    
                    # get the email body
                    if content_type == "text/plain":
                        body = msg.get_payload(decode=True).decode()
                        
                if content_type == "text/html":
                    if not os.path.isdir("attachments"):
                        # make a folder for this email (named after the subject)
                        os.mkdir("attachments")
                    filepath = os.path.join("attachments", "index.html")
                    # download attachment and save it
                    open(filepath, "w").write(body)
                    
                self.email_history.append({
                    "subject": subject,
                    "from": sender_address,
                    "body": body
                })
                
                self.new_emails.append({
                    "subject": subject,
                    "from": sender_address,
                    "body": body
                })

    async def updateClient(self, delay: int = 5):
        old_total_email_count = self.total_email_count
        
        # update the number of emails in the inbox
        status, messages = self.__imap.search(None, "ALL")
        self.total_email_count = len(messages[0].split())
        
        # calculate the number of new emails
        self.new_email_count = self.total_email_count - old_total_email_count
        
        # fetch the new emails
        for i in range(self.new_email_count):
            print("Fetching email " + str(self.total_email_count - i) + "...") # log the fetching of the email
            self.__fetch_email(self.total_email_count - i)
            
        await asyncio.sleep(delay) # wait for the next loop
    
    def resetNewEmails(self):
        self.new_emails = []
        self.new_email_count = 0

    def logoutAndClose(self):
        print("Logging out and closing email client...")
        self.__imap.close()
        self.__imap.logout()


class GenAIClient:
    def __init__(self, google_api_key: str, temperature: float = 0.9, top_p: float = 1, top_k: int = 1, max_output_tokens: int = 2048, system_instructions: str = "", model_name: str = "gemini-1.5-flash"):
        genai.configure(api_key=google_api_key)
        
        generation_config = genai.GenerationConfig(
            temperature=temperature,
            top_p=top_p,
            top_k=top_k,
            max_output_tokens=max_output_tokens
        )

        safety_settings = {
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
        }

        self.model = genai.GenerativeModel(model_name=model_name,
                                generation_config=generation_config,
                                safety_settings=safety_settings,
                                system_instruction=system_instructions)
        
        print("Starting chat with AI model...")
        
        self.chat = self.model.start_chat(history=[])
        
    async def parseNewMails(self, email_client: EmailClient, delay: int = 5):
        for i in range(email_client.new_email_count):
            email = email_client.new_emails[i]
        
            # send the email body to the AI model and save the response to a CSV file
            if email["body"] != "":
                print("Sending email " + str(i) + " to AI model...")
                response = self.chat.send_message(email["body"])
                data = json.loads(response.text)
                self.save_to_csv("email_" + str(i) + ".csv", data)
        
        email_client.resetNewEmails()
        
        await asyncio.sleep(delay) # wait for the next loop
        
    def save_to_csv(self, filename, data):
        path_name = "output"
        
        if not os.path.isdir(path_name):
            os.mkdir(path_name)
        
        print("Saving AI response to " + filename + "...")
        
        with open(os.path.join(path_name, filename), 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(data.keys())  # write header
            writer.writerow(data.values())  # write rows
    
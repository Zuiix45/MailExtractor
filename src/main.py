import os
import client
import gemini

USERNAME = os.getenv("EMAIL_ADDRESS")
PASSWORD = os.getenv("EMAIL_PASSWORD")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
OUTLOOK_IMAP = "outlook.office365.com"


if __name__ == "__main__":
    gemini.connectToGemini(GOOGLE_API_KEY)
    
    client.setTesseractPath("./tesseract/tesseract.exe")
    
    email_client = client.EmailClient()
    
    email_client.connectIMAP(OUTLOOK_IMAP)
    email_client.login(USERNAME, PASSWORD)
    email_client.selectMailbox("INBOX")
    email_client.setCurrentEmailIndex(1)
    email_client.fetchAndParse()
    email_client.logoutAndClose()
    
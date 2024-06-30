import os
import client
import gemini
import ui

USERNAME = os.getenv("EMAIL_ADDRESS")
PASSWORD = os.getenv("EMAIL_PASSWORD")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")



if __name__ == "__main__":
    
    ui.login.launch()
    
    gemini.connectToGemini(GOOGLE_API_KEY)
    
    client.setTesseractPath("./tesseract/tesseract.exe")
    

    """
    email_client.connectIMAP(OUTLOOK_IMAP)
    email_client.login(USERNAME, PASSWORD)
    email_client.selectMailbox("INBOX")
    email_client.fetchAndParse(19)
    email_client.logoutAndClose()
    """
    
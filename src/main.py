import os
import client
import asyncio

USERNAME = os.getenv("EMAIL_ADDRESS")
PASSWORD = os.getenv("EMAIL_PASSWORD")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
SYSTEM_INSTRUCTIONS = os.getenv("SYSTEM_INSTRUCTIONS")
OUTLOOK_IMAP = "outlook.office365.com"
        
if __name__ == "__main__":
    email_client = client.EmailClient(USERNAME, PASSWORD, OUTLOOK_IMAP, starting_email_index=0)
    genai_client = client.GenAIClient(GOOGLE_API_KEY, system_instructions=SYSTEM_INSTRUCTIONS)
    
    loop = asyncio.new_event_loop()
    
    try:
        asyncio.set_event_loop(loop)
        asyncio.ensure_future(email_client.updateClient())
        asyncio.ensure_future(genai_client.parseNewMails(email_client))
        loop.run_forever()
    except KeyboardInterrupt:
        print("Exiting...")
    finally:
        loop.close()
        email_client.logoutAndClose()
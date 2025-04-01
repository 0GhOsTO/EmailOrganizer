#pip install pywin32 in TERMINAL
#pip install pandas
#pip install openpyxl
import os
import win32com.client
import openai
import pandas as pd
import re
#TODO
#If the message was not - undeliverable(unable to identify), then categorize as green | Add into excel in border textform
#If the category marked as RED category, skip the message. 


#Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox folder
inbox = outlook.Folders("Sonia, SSW").Folders("Inbox")  # 6 corresponds to the Inbox folder
# Get all messages in the inbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", False)

#File path to Excel file
file_path = r"C:\Users\Andrew Cho\Desktop\Work\SoniaCopy.xlsx"
df = pd.read_excel(file_path)
new_bounced = []


# Loop through the emails
count =0
for message in messages:
    if(count == 400 ):
        break
    count +=1
    print("COUNT: ", count)
    # if message.Categories == "Red Category":
    #     print("RED Count: ", count)
    #     continue
    # else: 
    #     #Mark as blue category by default
    #     print("Marked as blue? ")
    #     message.Categories = "Blue Category"
    #     message.Save()

    #MOVE AROUND WITH THIS
    if message.Categories:
        print("Already categorized")
    else:
        message.Categories = "Blue Category"
        message.Save()


    # print("Count: ", count)

    try:
        sub = message.Subject
        messageUndeliverable = False
        print("SUB: ", message.Subject)
        keyword = "Undeliverable"

        #Check if the message's title contain undeliverable.
        if keyword.lower() in message.Subject.lower():
            messageUndeliverable = True
        # print("############", messageUndeliverable)
        # print(f"Subject: {message.Subject}")
        # print(f"Body: {message.Body}")

        #Look for the pattern with "recipents: blahblah"
        match = re.search(r"recipient:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", message.Body, re.IGNORECASE)

        #If they have it and message title undeliverable, then add as dictionary
        #else return recipant email is not found. 
        if match:
            if(messageUndeliverable == True):
                email = match.group(1)
                print("############Found recipient email:", match.group(1))
                if email not in df["Email"].values:
                    new_bounced.append({
                        "Email": email,
                        "Bounced Back": "Yes"
                    })
                    message.Categories = "Red Category"
                    message.Save()
                    print("RED Count: ", count)
                else:
                    print("Email already exists in Excel. Skipping.")
            messageUndeliverable = False
        else:
            print("############No recipient email found.")
        print("-" * 40)
    except Exception as e:
        print(f"####Error reading message: {e}")

rm_dup = []
seen = set()
for key in new_bounced: 
     em = key["Email"]
     if em not in seen:
          seen.add(em)
          rm_dup.append(key)

new_bounced = rm_dup


if new_bounced:
    updated_df = pd.concat([df, pd.DataFrame(new_bounced)], ignore_index=True)
    updated_df.to_excel(file_path, index=False)
    print(f"Added {len(new_bounced)} bounced emails to: {file_path}")
else:
    print("No new bounced emails found.")
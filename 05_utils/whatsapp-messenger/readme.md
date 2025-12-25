# WhatsApp Auto Messenger (Python + Selenium)

This project allows you to automatically send personalized WhatsApp messages to multiple contacts using Python, Selenium, and Chrome WebDriver.

---

## Prepare Your Files
### contacts.csv
Note: Use full mobile numbers with country code (without spaces, dashes, or plus sign).

    ```
    number,Student_Name,Naqeeb_Full_Name,Naqeeb_Contact_Number
    '+919701567398',Shaik Ishaq,Md Rizwan,+91 81790 23418
    '+918179724438',Syed Abdul Asif,Md Rizwan,+91 81790 23418
    ```
### message_template.txt
```
    Assalamu Alaikum {Student_Name},
    Hope you are doing well!
    Automated message from Ahsan Bhai
```    
## Running the Script
```bash 
python send_whatsapp_messages_with_selenium.py
```
- WhatsApp Web will open automatically.
- Scan your QR code using your phone; then press enter on the running program
- Messages will start sending automatically.

## Important Notes
- Keep your phone connected to the internet throughout the process.
- Do not close the Chrome window during message sending.
- It's recommended not to send too many messages at once to avoid Whatapps(Meta) blocking your number.

## Applied to Naqeeb Sheet to generate data neeed in contacts.csv
- Formula to extract whatsapp number in `+01 <number>` format. Formula: `="'" & "+91" & TEXT(E4, "0") & "'"`, call this column 'WhatsApp# 2'. Assume its in 'AM4' cell. 
-  Then add Naqeeb Number on column 'AL4'. 
-  Then on column AN4 apply `=AM4 & "," & C4 & "," & AK4 & "," & AL4`. This will generate output as `'+919701567398',Shaik Ishaq,Md Rizwan,+91 81790 23418`; 
- Paste these contains on contacts.csv file and run `send_whatsapp_messages_with_selenium.py`

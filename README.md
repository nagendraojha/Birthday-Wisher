
# Auto Birthday Wisher


---

### 🔹 **1. Importing Libraries**
```python
import pandas as pd
import datetime
import smtplib
import os
```
- `pandas`: For reading and updating the Excel file.
- `datetime`: To get the current date and time.
- `smtplib`: To send emails via Gmail's SMTP server.
- `os`: To work with the file system (like getting current working directory).

---

### 🔹 **2. Getting Current Directory**
```python
current_path = os.getcwd()
print("Current path:", current_path)
os.chdir(current_path)
```
- Prints the path where the script is running.
- Ensures the script is reading/writing files in the correct directory.

---

### 🔹 **3. Email Login**
```python
GMAIL_ID = input("Enter your email: ").strip()
GMAIL_PSWD = input("Enter your app password: ").strip()
```
- Takes the sender's Gmail ID and **App Password** (not your main password!).
- `.strip()` removes extra spaces.

---

### 🔹 **4. Function to Send Email**
```python
def sendEmail(to, sub, msg):
    ...
```
- Sends an email using the Gmail SMTP server.
- Uses `TLS` encryption (port 587).
- Converts the message to UTF-8 to support emojis and special characters.
- Catches and prints any errors during sending.

---

### 🔹 **5. Main Logic**
```python
if __name__ == "__main__":
```
This is the core logic that checks birthdays and sends emails.

---

### 🔸 Load the Excel File
```python
df = pd.read_excel("data.xlsx")
```
Reads your birthday data from an Excel file into a DataFrame.

---

### 🔸 Get Current Time
```python
now = datetime.datetime.now()
today_str = now.strftime("%d-%m")
yearNow = now.strftime("%Y")
```
- `now`: full timestamp.
- `today_str`: current date (e.g., "15-04").
- `yearNow`: current year.

---

### 🔸 Loop Through People
```python
for index, item in df.iterrows():
```
This loop checks each row in your Excel file:

#### Inside the loop:
```python
raw_bday = str(item['Birthday']).strip()
bday = datetime.datetime.strptime(raw_bday, "%d-%m-%Y").strftime("%d-%m")
```
- Converts the birthday to `"dd-mm"` format for matching with today.

```python
last_wished_at = item.get('LastWishedAt', '')
```
- Tries to get the last time a birthday wish was sent.

#### Logic to Avoid Duplicate Wishes Within 24 Hours:
```python
if pd.notna(last_wished_at):
    last_wished_dt = pd.to_datetime(last_wished_at)
    diff = now - last_wished_dt
    if diff.total_seconds() < 86400:
        can_send = False
```
If the person was already wished in the last 24 hours, skip sending again.

---

### 🔸 Send Email if Conditions Met
```python
if today_str == bday and can_send:
    sendEmail(email, "🎉 Happy Birthday!", message)
```
If it’s their birthday **and** they haven’t been wished today → send the email.

---

### 🔸 Update Excel with Time of Last Wish
```python
df.loc[i, 'LastWishedAt'] = now.strftime("%Y-%m-%d %H:%M:%S")
```
Stores the timestamp when a wish was sent to avoid duplicates within 24 hours.

---

### 🔸 Save Changes
```python
df.to_excel('data.xlsx', index=False)
```
Saves the updated DataFrame (with new `LastWishedAt` values) back to the Excel file.

---

### ✅ Summary
This script:
- Loads a birthday list from Excel.
- Checks if today is anyone’s birthday.
- Makes sure you haven’t already wished them in the past 24 hours.
- Sends a 🎉 birthday email using Gmail.
- Updates the Excel to track when the wish was sent.



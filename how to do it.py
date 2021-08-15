Challenge: Every day, our system updates sales from the previous day. Your daily job, as an analyst, is to send an email to the board of directors, as soon as you start work, with the billing and quantity of products sold the day before.

Board of directors email: blablabla+directors@gmail.com

Location where the system makes the sales of the previous day available: https://drive.google.com/drive/folders/blablablabla

To solve this, let's use pyautogui, a keyboard and mouse command automation library.

# this was done using the Jupyter Notebook

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Step 1 - Enter inside our system (https://drive.google.com/drive/folders/blablablabla)
pyautogui.hotkey('ctrl', 't')
link = 'https://drive.google.com/drive/folders/blablablabla'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# Step 2 - Navigate inside the system (enter inside a file)
time.sleep(5) # wait 5 seconds
pyautogui.click(x=-1535, y=354, clicks=2)

# Step 3 - Download the file of sales
time.sleep(2)
pyautogui.click(x=-1516, y=446) # click on the file
pyautogui.click(x=-206, y=191) # click on options
pyautogui.click(x=-391, y=594) # click on download
time.sleep(3) # wait for the download

# Step 4 - Calculate billing and quantity of products sold
import pandas as pd

chart = pd.read_excel(r"C:\Users\MyPC\Desktop\File\Sales - Dez.xlsx")
billing = chart["Final Value"].sum()
quantity = chart["Quantity"].sum()
display(chart)
print(billing)
print(quantity)

# Step 5 - Send the email to the board of directors
pyautogui.hotkey('ctrl', 't') # it opens a new tab
link = 'https://mail.google.com/'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# click on write button
time.sleep(3)
pyautogui.click(x=-1808, y=243)

# write to the email that I'm sending the email
time.sleep(3)
pyautogui.write('blablabla+directors@gmail.com')
pyautogui.press('tab') # choose the email
pyautogui.press('tab') # go to the subject field

# write the subject
subject = "Sales Report"
pyperclip.copy(subject)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab') # go to the email field

# write the email
text = f"""
Dears, good morning

The billing was R${billing:,.2f}
The quantity of products was {quantity:,.2f}

Thanks for the attention
Bruno Henrique"""
pyautogui.write(texto)

# send the email

pyautogui.hotkey('ctrl', 'enter')

# how to get the position of the mouse cursor
import time
time.sleep(5)
pyautogui.position()
Point(x=-1277, y=273) #example

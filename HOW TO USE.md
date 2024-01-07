# How to Use Bas.AI
## Clone the repository
```bash
git clone https://github.com/atharv-rem/Bas.AI.git
```
## Word-Bas.AI [![Word](https://img.shields.io/badge/Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white)](https://your-link-here)
---
Add your openAI API key <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L22-L22)</kbd>
```python
openai.api_key = "" #Add your own OpenAI key
```

Add file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L304)</kbd>
```python
plt.savefig(f'', dpi=600, bbox_inches='tight',
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add the same file path <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L307)</kbd>
```python
doc.add_picture(f'', width=Cm(16.51),
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L539)</kbd>
```python
plt.savefig(f'', dpi=600, bbox_inches='tight',
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph2}.png
```
Add the same file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L542)</kbd>
```python
doc.add_picture(f'', width=Cm(16.51),
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path to save the word file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L628)</kbd>
```python
file_name = f''  # file path in this format (keep this variable)C:\\Users\\OneDrive\\Desktop\\code\\{H1} - Annual Performance.docx
```
## Excel-Bas.AI [![Excel](https://img.shields.io/badge/Excel-008000?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://your-link-here)
---
Add file path to save the excel file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Excel.py#L152C7-L152C7)</kbd>
```python
file_name = f'' #file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{Name} - Annual Performance.xlsx
```
## PPT-Bas.AI [![PowerPoint](https://img.shields.io/badge/PowerPoint-FF4F28?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)](https://your-link-here)
---
Add file path to save the word file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/PPT.py#L292C1-L293C117)</kbd>
```python
prs.save(f"") # file path in this format(keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{file_name}
os.startfile(f"")  # file path in this format(keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{file_name}
```
## Email-Bas.AI [![Gmail](https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white)](https://your-link-here)
---
Fill these out <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L11C1-L22C31)</kbd>
```python
# Setup port number and server name
smtp_port = 587  # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server
# OpenAI key to access OpenAI account
openai.api_key = "" #your API key
# Password to access Gmail account
pswd = '' #Password to your gmail account
# Email id from which the mails have to be sent
email_from = "" #Your Email ID
```
**You can get the password to your gmail account by following the steps:-**
- Go to **Manage your Google account**
- under **security** click **2step verification**
- Scroll down and go to **App passwords**
- create new and give the name as **python**
- copy the password and paste it in the **pswd** variable

**You can get the API key of your API account by following the steps:-**
- Login to your account
- Go to **API**
- Go to the menu on the top right
- click **API keys**
- create new key and paste it into **openai_key** variable
  
Add directory path <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L166C8-L166C8)</kbd> 
```python
for i in os.listdir(r''):  # add your own directory from which the files are being displayed
```
Add the file path related to your directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L229C20-L229C20)</kbd>
```python
filename = f'' # file path in this format (Keeep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{v}
```
## GUI-Bas.AI [![UI Design](https://img.shields.io/badge/UI%20Design-EAA52C?style=for-the-badge&labelColor=purple)](https://your-ui-design-link-here)
---
Add the file path to your word python code <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/3875b93b12d9be55a5d57f97656f89c22bd7189d/GUI.py#L47)</kbd>
```python
subprocess.Popen(["python", ""]) #Add your Word-Bas.AI path
```
Add the file path to your excel python code <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/3875b93b12d9be55a5d57f97656f89c22bd7189d/GUI.py#L55)</kbd>
```python
subprocess.Popen(["python", ""]) #Add your Excel-Bas.AI path
```
Add the file path to your ppt python code <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/3875b93b12d9be55a5d57f97656f89c22bd7189d/GUI.py#L62)</kbd>
```python
subprocess.Popen(["python",""]) #Add your PPT-Bas.AI path
```
Add the file path to your mail python code <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/3875b93b12d9be55a5d57f97656f89c22bd7189d/GUI.py#L71)</kbd>
```python
subprocess.Popen(["python", ""]) #Add your Send_Mail-Bas.AI path
```

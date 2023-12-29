# How to Guide
## Step1 - Clone the repository
```bash
git clone https://github.com/atharv-rem/Bas.AI.git
```
## Word-Bas.AI [![Word](https://img.shields.io/badge/Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white)](https://your-link-here)
Add your openAI API key <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L22-L22)</kbd>
``` 
openai.api_key = "" #Add your own OpenAI key
```

Add file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L304)</kbd>
```
plt.savefig(f'', dpi=600, bbox_inches='tight',
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add the same file path <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L307)</kbd>
```
doc.add_picture(f'', width=Cm(16.51),
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L539)</kbd>
```
plt.savefig(f'', dpi=600, bbox_inches='tight',
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph2}.png
```
Add the same file path of where the graph has to be stored <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L542)</kbd>
```
doc.add_picture(f'', width=Cm(16.51),
#add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path to save the word file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L628)</kbd>
```
file_name = f''  # file path in this format (keep this variable)C:\\Users\\OneDrive\\Desktop\\code\\{H1} - Annual Performance.docx
```
## Excel-Bas.AI [![Excel](https://img.shields.io/badge/Excel-008000?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://your-link-here)
Add file path to save the excel file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Excel.py#L152C7-L152C7)</kbd>
```
file_name = f'' #file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{Name} - Annual Performance.xlsx
```
## PPT-Bas.AI [![PowerPoint](https://img.shields.io/badge/PowerPoint-FF4F28?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)](https://your-link-here)
Add file path to save the word file in your required directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/PPT.py#L292C1-L293C117)</kbd>
```
prs.save(f"") # file path in this format(keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{file_name}
os.startfile(f"")  # file path in this format(keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{file_name}
```
## Email-Bas.AI [![Gmail](https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white)](https://your-link-here)
Fill these out <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L11C1-L22C31)</kbd>
```
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
Add directory path <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L166C8-L166C8)</kbd> 
```
for i in os.listdir(r''):  # add your own directory from which the files are being displayed
```
Add the file path related to your directory <kbd>[here](https://github.com/atharv-rem/Bas.AI/blob/b7cbe3381b8621b73f9e8d36605f70cb4813a3fc/Send%20mail.py#L229C20-L229C20)</kbd>
```
filename = f'' # file path in this format (Keeep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{v}
```
## GUI-Bas.AI [![UI Design](https://img.shields.io/badge/UI%20Design-EAA52C?style=for-the-badge&labelColor=purple)](https://your-ui-design-link-here)
for GUI to work all the above python code files have to be in the same folder as the GUI.

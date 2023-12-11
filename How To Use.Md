# How to Guide
## Step1 - Clone the repository
```bash
git clone https://github.com/atharv-rem/Bas.AI.git
```
## Word-Bas.AI [![Word](https://img.shields.io/badge/Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white)](https://your-link-here)
Add your openAI API key [here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L22-L22)
``` openai.api_key = "" #Add your own OpenAI key
```

Add file path of where the graph has to be stored [here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L304)
```plt.savefig(f'', dpi=600, bbox_inches='tight',  #add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add the same file path [here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L307)
```doc.add_picture(f'', width=Cm(16.51),  #add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path of where the graph has to be stored[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L539)
```plt.savefig(f'', dpi=600, bbox_inches='tight', #add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph2}.png
```
Add the same file path of where the graph has to be stored[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L542)
```doc.add_picture(f'', width=Cm(16.51), #add file path in this format (keep the variable) C:\\Users\\OneDrive\\Desktop\\code\\{TofGraph1}.png
```
Add file path to save the word file in your required directory[here](https://github.com/atharv-rem/Bas.AI/blob/7274ac588e537616122b36a0247a4eb836d3c7ff/Word.py#L628)
```file_name = f''  # file path in this format (keep this variable) C:\\Users\\OneDrive\\Desktop\\code\\{H1} - Annual Performance.docx
```



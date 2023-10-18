Quick test on how to generate some synthetic data with ChatGPT. Still improvements are nesseccary:
- ChatGPT responses are sometimes unrelyable in their consistency -> needs more cleaning of data and workarounds.
- Better automation of texts.
- More finetuning.
- Clean up.
Used for a study project.

# How to generate some data
### Requirements
Account at openai, go to tab API keys, generate API key, put it in file.
```
pip install openai
pip install reportlab
```
### Generate Files
The current project creation date is from last year, this goes for all projects except for RuhrEnergieSolutions! If you want to create Data for RuhrEnergieSolutions you have to switch the creation date in the main to the one from this year.
If you notice that few files are generated, because chatgpt ignores the template, and you want more, change the numbers in createfolder() and createSubfolder ().
Topics and Department should match!
Start generation with:

```
Linux
python3 openaiFileGen_linux.py --mainfolder /home/username/where/you/want  --department IT-Department --topic 1 --project 0 --number 4

optional arguments:
  -h, --help            show this help message and exit
  --mainfolder MAINFOLDER
                        path to the credentials file
  --department {IT-Department,HR-Department,Management-Department,PR-Depatrment,Projekt-Department}
                        Select a department ["IT-Department", "HR-Department", "Management-Department", "PR-Depatrment", "Projekt-Department"]
  --project {0,1,2,3,4}, -p {0,1,2,3,4}
                        ["NEXUSNetCity", "RuhrEnergieSolutions", "LaunchPad", "UrbanPlanT", "TransportPlane"]
  --topic {0,1,2,3}, -t {0,1,2,3}
                        Pick a number, starts with 0, [topicsHR, topicsIT, topicsPM, topicsPR, topicsManagement]
  --number NUMBER, -n NUMBER
                        number of topfolders
```

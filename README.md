# Summary
This project aims to entail a few tools that help automating content creation within Arm Education and Research Enablement.

This project aims to entail a few tools that help automating content creation within Arm Education and Research Enablement. The tool can be found under the "PPT" folder and navigate to the pptx windows (Please be aware that the tool will only run in Windows). It contains three python script Refactor.py, MapPre.py, TestTool.py. The main script that needs to run is the Refactor.py.

More documentation to follow.

visit https://confluence.arm.com/display/RSHEDU/EduSuite+Wiki#users-statistic to get a better understanding of the project

visit http://marmor05.p.research.arm.com/edusuitehtmldoc/index.html to understand the content of each module

This project makes use of the PPTX python module.

## Installation
Before running the script make sure if you have Python3. To run the script there are some python dependencies that needs to be installed

Use the package manager [pip3](https://pip.pypa.io/en/stable/) to install dependencies.

```bash
pip3 install requirements.txt
```

## Usage

*Notes* : Do not have any open powerpoint files while running the tool. Disable or pause OneDrive or similar backup apps, which might interfere with the tool running. Template provided should have no slides, just the slide master.

After installing all the above python packages, run the script Refactor.py,
```bash
python  Refactor.py
```
The program will ask to enter whether you want to specify a custom directory or default directory (the one where the program is located) to extract lectures, templates and create a new folder where the new version of the lectures will be stored.
Furthermore, using a command input, the user will be able to:

* change and select a template
* exit the program
* map all the presentations that were found in the directory
* map a specific presentation
* When mapping a presentation, the new template will be applied to the selected presentation

## PPT folder structure
* pptx_linuxV1 folder: In this folder you will find  first prototype of the PPT tool
  * dogs folder: a folder containing images of dogs
  * template folder: a folder containing templates examples for mapping lectures
  * mappingPresentationEdu.py: first version of the method for mapping presentation using https://python-pptx.readthedocs.io/en/latest/
  * scanLecturesV2.py: first version of script to look for templates in a directory and find presentations and templates
* pptx_linuxV2 folder contains an improved version of mappingPresentationEdu.py and a prototype method (xdxd.py) for mapping presentions(only works on windows) using win32 https://pypi.org/project/pywin32/
* __pptx_windows folder contain the latest development of the tool__, this should be the folder where we should continue to develop
  * templatesForGG folder: some default templates that can be used when running the tool
  * MapPre.py: this script contains a method that applies a new template to an old presentation
  * Refactor.py: this script in charged of everything
  * TestTool.py: this is a file containing unittesting using the unnitest framework https://docs.python.org/3/library/unittest.html
  * Full documentation of these files can be found in the following link: http://marmor05.p.research.arm.com/edusuitehtmldoc/index.html
* release folder: this folder should be used to store build version of the PPT tool like executables(.exe)
## To do list
All the current EduSuite issues can be found here → https://jira.arm.com/browse/AEMCC-312
* Disable pdf generation by default. Give user the option to trigger this option using the command prompt
* Files with the same file names need to be recoginized by its path
* Command line feature, specify directories as arguments
* Warning messages
* Writing Integration test for EduSuite PPT

Some suggestion on how to solve these issue have been added in the confluence page → https://confluence.arm.com/display/RSHEDU/EduSuite+Wiki
## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

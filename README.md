
# PowerPoint Branding Tool

Welcome to the PowerPoint Branding Tool.

### [Download the tool here](https://github.com/arm-university/PPT-branding-tool/archive/refs/heads/main.zip)

This project makes it easy to change branding and template styles across multiple PowerPoint files. It may be useful alongside our Education Kits, for partners and Academics who wish to apply their own branding and styles to our materials. The tool can be found under the "PPT" folder and navigate to the pptx windows (Please be aware that the tool will only run in Windows). It contains three python script Refactor.py, MapPre.py, TestTool.py. The main script that needs to run is the Refactor.py.

## Getting Involved
We welcome contributions, amendments & modifications to this tool. For details, please click on the following links:

* [How to contribute](https://github.com/arm-university/PPT-branding-tool/blob/main/Contributions_and_Modifications/Contributions_and_Modifications.md)
* [Type of modifications](https://github.com/arm-university/PPT-branding-tool/blob/main/Contributions_and_Modifications/Desired_Contributions.md) we are looking for. We also use [Projects](https://github.com/arm-university/PPT-branding-tool/projects) to track progress.
* [Workflow](https://github.com/arm-university/PPT-branding-tool/blob/main/Contributions_and_Modifications/workflow.pdf)


## License
You are free to amend, modify, fork or clone this material. See [License.md](https://github.com/arm-university/PPT-branding-tool/blob/main/License/License.md) for the complete license.

## Inclusive Language Commitment
Arm is committed to making the language we use inclusive, meaningful, and respectful. Our goal is to remove and replace non-inclusive language from our vocabulary to reflect our values and represent our global ecosystem.
 
Arm is working actively with our partners, standards bodies, and the wider ecosystem to adopt a consistent approach to the use of inclusive language and to eradicate and replace offensive terms. We recognise that this will take time. This course contains references to non-inclusive language; it will be updated with newer terms as those terms are agreed and ratified with the wider community. 
 
Contact us at education@arm.com with questions or comments about this tool. You can also report non-inclusive and offensive terminology usage in Arm content at terms@arm.com.

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
* pptx_linuxV1 folder: In this folder you will find  first prototype of the PowerPoint branding tool
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
* release folder: this folder should be used to store build version of the PowerPoint branding tool like executables(.exe)

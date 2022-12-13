
# PPT Branding Tool

Welcome to the PPT Branding Tool.

### [Download the tool here](https://github.com/arm-university/PPT-branding-tool/archive/refs/heads/main.zip)

Our flagship offering to universities worldwide is the Arm University Program Education Kit series.

These self-contained educational materials offered exclusively and at no cost to academics and teaching staff worldwide. They’re designed to support your day-to-day teaching on core electronic engineering and computer science subjects. You have the freedom to choose which modules to teach – you can use all the modules in the Education Kit or only those that are most appropriate to your teaching outcomes.

Our Rapid Embedded Systems Design Education Kit covers the fundamental principles of how to accelerate the development of embedded systems and rapidly prototype various embedded applications. A full description of the education kit can be found [here](https://www.arm.com/resources/education/education-kits/rapid-embedded-systems). 

## Getting Involved
We welcome contributions, amendments & modifications to this education kit. For details, please click on the following links:

* [How to contribute](https://github.com/arm-university/Rapid-Embedded-Education-Kit/blob/main/Contributions_and_Modifications/Contributions_And_Modifications.md)
* [Type of modifications](https://github.com/arm-university/Rapid-Embedded-Education-Kit/blob/main/Contributions_and_Modifications/Desired_Contributions.md) we are looking for. We also use [Projects](https://github.com/arm-university/Rapid-Embedded-Education-Kit/projects) to track progress.
* [Workflow](https://github.com/arm-university/Rapid-Embedded-Education-Kit/blob/main/Contributions_and_Modifications/workflow.pdf)


 ## Kit specification:

* A full set of lecture slides, ready for use in a typical 10-12-week undergraduate course (full syllabus below) .
* Lab manual with code solutions for faculty. Labs are based on low-cost but powerful Arm-based hardware platforms. 
* **Prerequisites:** Basics of programming in C / C++.

## Course Aim
To produce students who can design and program Arm-based embedded systems and implement them using commercial API.

## Syllabus
1. Introduction to Embedded Systems
1. The Arm Cortex-M4 Processor Architecture
1. Introduction to Arm Cortex-M4 Programming
1. Introduction to the Mbed Platform and CMSIS
1. Digital Input and Output (IO)
1. Interrupts and Low Power Features
1. Analog Input and Output
1. Timer and Pulse-Width Modulation
1. Serial Communication
1. Real-Time Operating Systems
1. Final Project: Audio Player

**Extra Reading:** The Arm Cortex-M Processor Architecture: Part 2.

## License
You are free to amend, modify, fork or clone this material. See [LICENSE.md](https://github.com/arm-university/Rapid-Embedded-Education-Kit/blob/main/License/LICENSE.md) for the complete license.

## Inclusive Language Commitment
Arm is committed to making the language we use inclusive, meaningful, and respectful. Our goal is to remove and replace non-inclusive language from our vocabulary to reflect our values and represent our global ecosystem.
 
Arm is working actively with our partners, standards bodies, and the wider ecosystem to adopt a consistent approach to the use of inclusive language and to eradicate and replace offensive terms. We recognise that this will take time. This course contains references to non-inclusive language; it will be updated with newer terms as those terms are agreed and ratified with the wider community. 
 
Contact us at education@arm.com with questions or comments about this course. You can also report non-inclusive and offensive terminology usage in Arm content at terms@arm.com.

------------------------------

In this material, we use the terms ‘Controller’ and 'Target' in the context of the the I2C and SPI protocols, instead of the terms ‘Master’ and ‘Slave’, which were conventionally used until recently. As a result, the related concepts of 'MISO' and 'MOSI' become 'CITO' and 'COTI'.









# PPT Branding Tool
This project helps automate content creation within Arm Education and Research Enablement. The tool can be found under the "PPT" folder and navigate to the pptx windows (Please be aware that the tool will only run in Windows). It contains three python script Refactor.py, MapPre.py, TestTool.py. The main script that needs to run is the Refactor.py.

## Arm Internal
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

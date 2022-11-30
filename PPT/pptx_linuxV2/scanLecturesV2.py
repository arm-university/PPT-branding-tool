from pptx import Presentation
import mappingPresentationEdu as mapPt
import os
import shutil
import sys
import xdxd
#https://stackoverflow.com/questions/3964681/find-all-files-in-a-directory-with-extension-txt-in-python
pathFile = None
filesAndPathFileTemplate = {}

#currently this program will extract all the pptx from the currently dir
#transversing through the folder and picking templates and presentation
#that DON'T contain tagForRecognisingNewPresentations
user_input = input("Enter a path to execute this program: ")

assert os.path.exists(user_input), "The following path does not exist:"+str(user_input)

dir = user_input
filesAndPathFilesPresentation = {}
tagForRecognisingNewPresentations = "newVersion"
for root, dirs, files in os.walk(dir):
    for file in files:
        if file.endswith(".pptx"):
             if "template" in file:
               pathFile = os.path.join(root, file)
               filesAndPathFileTemplate[file] = pathFile
             elif tagForRecognisingNewPresentations not in file:
               pathFile = os.path.join(root, file)
               filesAndPathFilesPresentation[file] = pathFile


#to print results in order
filesAndPathFileSortedTemplate = sorted(filesAndPathFileTemplate)
filesAndPathFilesSortedPresentation = sorted(filesAndPathFilesPresentation)
mappingNumberToKeyDicTemplate = []
mappingNumberToKeyDicPresentation = []

for i in filesAndPathFileSortedTemplate:
    mappingNumberToKeyDicTemplate.append(i)
for i in range(len(mappingNumberToKeyDicTemplate)):
    print(str(i) + " = [" + mappingNumberToKeyDicTemplate[i] + "]")

templateName = None
numberForTemplate = input("\nEnter the number of the template to be used: ")

try:
  numTemplate = int(numberForTemplate)
  if (numTemplate >= len(mappingNumberToKeyDicTemplate) or numTemplate < 0):
      raise ValueError()
except ValueError:
  print(mapPt.colored('\nERROR: ', 'red') + "Please input an integer between the range of 0 - " + str(len(mappingNumberToKeyDicTemplate) - 1))
  sys.exit()

templateNameKey = mappingNumberToKeyDicTemplate[int(numTemplate)]

for i in filesAndPathFilesSortedPresentation:
    mappingNumberToKeyDicPresentation.append(i)

dir = dir + "/newSlides"
if not os.path.exists(dir):
  os.makedirs(dir)

presenationVisited = len(mappingNumberToKeyDicPresentation) - 1
isChangingTemplate = False

while(True):
    if(presenationVisited == len(mappingNumberToKeyDicPresentation) - 1):
        letter = input("[q] for exit \n[a] for mapping all presentations in the current directory \n[c] for changing the template \nAny other letter to select a specific presentation \n\nEnter your selection: ").lower()
        if(letter == "q"):
          break
        elif(letter == "a"):
          presenationVisited = 0
          presentationChosen = presenationVisited
        elif(letter == "c"):
          isChangingTemplate = True
          for i in range(len(mappingNumberToKeyDicTemplate)):
              print(str(i) + " = [" + mappingNumberToKeyDicTemplate[i] + "]")
          templateChosen = input("\nEnter the number of the template you would like to use: ")
          try:
            templateNumberChose = int(templateChosen)
            if (templateNumberChose >= len(mappingNumberToKeyDicTemplate) or templateNumberChose < 0):
                raise ValueError()
          except ValueError:
            print(mapPt.colored('\nERROR: ', 'red') + "Please input an integer between the range of 0 - " + str(len(mappingNumberToKeyDicTemplate) - 1))
            sys.exit()
          templateNameKey = mappingNumberToKeyDicTemplate[templateNumberChose]
        else:
            for i in range(len(mappingNumberToKeyDicPresentation)):
                print(str(i) + " = [" + mappingNumberToKeyDicPresentation[i] + "]")
            presentationChosen = input( "\nEnter the number of the presentation you would like to use: ")
            try:
              checkNumber = int(presentationChosen)
              if (checkNumber >= len(mappingNumberToKeyDicPresentation) or checkNumber < 0):
                  raise ValueError()
            except ValueError:
              print(mapPt.colored('\nERROR: ', 'red') + "Please input an integer between the range of 0 - " + str(len(mappingNumberToKeyDicPresentation) - 1))
              sys.exit()
    else:
      presenationVisited += 1
      presentationChosen = presenationVisited

    if(isChangingTemplate == False):
        #oldtemplateSelected = Presentation(filesAndPathFileTemplate[templateNameKey])
        mewTemplateSelected = filesAndPathFileTemplate[templateNameKey]

        mappingkey = mappingNumberToKeyDicPresentation[int(presentationChosen)]
        #mapPt.new_temp = Presentation(filesAndPathFilesPresentation[mappingkey])
        oldPresentationSelected = filesAndPathFilesPresentation[mappingkey]
        #deletingextension = mappingNumberToKeyDicTemplate[int(presentationChosen)].split(".pptx")
        #resultMapping,errorsDetection  = mapPt.mappingPre(oldPresentationSelected, mewTemplateSelected)
        nameModification = dir + "/" + tagForRecognisingNewPresentations+ "-" + mappingkey
        xdxd.mappingPreV2(oldPresentationSelected,mewTemplateSelected,nameModification,dir)
        #
        #
        # prs2ToSave.SaveAs(nameModification)
        # if(errorsDetection == True):
        #     print(mapPt.colored('There was a problem with your mapping', 'red'))
        #     input("Before continuing ,add the missing slides manually and press Enter to resume this program\n")
        # prs2ToSave.Close()

    else:
        isChangingTemplate = False
        print("The template has been changed to: " +  templateNameKey)

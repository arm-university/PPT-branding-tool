"""
This confidential and proprietary software may be used only as
authorised by a licensing agreement from ARM Limited
(C) COPYRIGHT 2020 ARM Limited
ALL RIGHTS RESERVED
The entire notice above must be reproduced on all authorised
copies and copies may only be made to the extent permitted
by a licensing agreement from ARM Limited.
Author: Education Team
Date: 12/08/2020
Summary: This program will ask users to input rootpaths to
         extract all PowerPoint presentations into templates and lectures,
         one of the roots will be used to save the output of the program
         Furthermore, using a command input, the user will be able to:
         -change and select a template
         -exit the program
         -select to save presentations as .pptx or .pptx and .pdf files
         -map all the presentations that were found in the current directory
         -map a specific presentation
         When mapping a presentation, the new template will be applied in the
         presentation desired. This outputs a new PowerPoint document and an
         optional PDF.
         PLEASE CLOSE ANY PowerPoint process in your machine before running
         this program
"""
import os
import sys
import MapPre
import natsort

path_file = None
dict_templates = {}
dict_lectures = {}


def find_powerpoint_files(dir):
    """This function takes a directory and transverse through it to
    extract and store pptx files in a dictionary

    :param string dir: an existent path to a directory that
                       should contain PowerPoint presentations

    :return: a dictionary with the pptx files found in the directory
    :rtype: dictionary

    :raises Assertion Exception: if no pptx files are found
                                 in the directory
    """
    dict_temp = {}
    path_to_file = ""

    # windows could have hidden temporary PowerPoint files
    # that could cause erros, we will not add them to the dictionary
    # they start with the extension ~$
    hidden_temporary_files_ext = "~$"

    # create empty list to detect duplicated filekeys
    detected_filekeys = []

    # find the PPT files in all folders and subfolders of current directory
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith(".pptx") and not file.startswith(
                hidden_temporary_files_ext
            ):
                path_to_file = os.path.join(root, file)

                # make sure all keys are unique by adding relative paths
                filekey = os.path.normpath(os.path.relpath(path_to_file, dir))
                detected_filekeys.append(filekey)
                dict_temp[filekey] = path_to_file

    # check if there were duplicated filekey names somehow
    try:
        duplicate_file_list = list(
            {
                duplicates
                for duplicates in detected_filekeys
                if detected_filekeys.count(duplicates) > 1
            }
        )
        if len(duplicate_file_list):
            raise ValueError()
    except ValueError:
        print(
            MapPre.colored("\nERROR: ", "red")
            + "Duplicated filekeys in dictionary detected "
            + "for files "
            + str(duplicate_file_list)
            + ".\nPlease rerun program with manual selection of "
            + "paths. If problem persists, contact owner of the script."
        )

    # check if dictionary is empty
    try:
        if not bool(dict_temp):
            raise ValueError()
    except ValueError:
        print(
            MapPre.colored("\nERROR: ", "red")
            + "no pptx files were found while transversing the directory"
            + ", please try another directory"
        )
        sys.exit(1)
    return dict_temp


def find_powerpoint_lectures_and_templates(dir, tag):
    """This function takes a directory and transverse through it to
    extract and store pptx lectures and templates in a dictionary

    :param string dir: an existent path to a directory that
                       should contain PowerPoint presentations

    :param string tag: a tag/word/phrase that templates presentations
                       will contain in order to differentiate them from
                       lectures

    :return: a tuple containing two dictionaries, one of lectures
             and the another one of templates
    :rtype: tuple of dictionaries

    :raises Assertion Exception: if no pptx files are found
                                 in the directory

    """
    dict_lect_temp = {}
    dict_templates_temp = {}
    path_to_file = ""
    # windows could have hidden temporary PowerPoint files
    # that could cause erros, we will not add them to the dictionaries
    # they start with the extension ~$
    hidden_temporary_files_ext = "~$"

    # create empty list to detect duplicated filekeys
    detected_filekeys = []

    # find the PPT files in all folders and subfolders of current directory
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith(".pptx") and not file.startswith(
                hidden_temporary_files_ext
            ):
                path_to_file = os.path.join(root, file)

                # make sure all keys are unique by adding relative paths
                filekey = os.path.normpath(os.path.relpath(path_to_file, dir))
                detected_filekeys.append(filekey)

                if tag in file:
                    dict_templates_temp[filekey] = path_to_file
                else:
                    dict_lect_temp[filekey] = path_to_file

    # check if there were duplicated filekey names somehow
    try:
        duplicate_file_list = list(
            {
                duplicates
                for duplicates in detected_filekeys
                if detected_filekeys.count(duplicates) > 1
            }
        )
        if len(duplicate_file_list):
            raise ValueError()
    except ValueError:
        print(
            MapPre.colored("\nERROR: ", "red")
            + "Duplicated filekeys in dictionary detected "
            + "for files "
            + str(duplicate_file_list)
            + ".\nPlease rerun program with manual selection of "
            + "paths. If problem persists, contact owner of the script."
        )

    # check if either dictionary is empty
    try:
        is_dir_lect_empty = bool(dict_lect_temp)
        is_dir_template_empty = bool(dict_templates_temp)
        if not (is_dir_lect_empty and is_dir_template_empty):
            raise ValueError()
    except ValueError:
        pptx_files_not_found = ""
        if not is_dir_lect_empty:
            pptx_files_not_found += "[lectures]"
        if not is_dir_template_empty:
            pptx_files_not_found += "[templates]"
        print(
            MapPre.colored("\nERROR: ", "red")
            + "No pptx files were found while transversing the directory for:"
            + " "
            + pptx_files_not_found
            + ", please try another directory"
        )
        sys.exit(1)
    return (dict_lect_temp, dict_templates_temp)


def general_get_input(text):
    """This function works as a wrapper in order to test other functions

    :param string text: it represents a default message before input

    :return: a string returned from the input() method
    :rtype: string
    """
    return input(text)


def make_new_dir(dir_path):
    """This function create a new directory

    :param string dir_path: path where we will create the new dir
    """
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


def user_input_path(msg_for_user):
    """This function asks the user to insert
    a path and check whether it exists in the system

    :param string msg_for_user: a text will be sent to general_get_input

    :return: an existent path in the system
    :rtype: string

    :raises Assertion Exception: if path does not exist
    """
    user_input = general_get_input(msg_for_user)
    assert os.path.exists(user_input), "The following path does not exist:" \
        + str(user_input)
    return user_input


def select_a_number(mapping_num_to_key_dict, item):
    """This function takes a list and map its elements with a number
    , then the user can choose what element he would like to extract
    from the list

    :param list mapping_num_to_key_dict: a list with the key from the
                                         the templates or presentations
                                         dictionaries
    :param string item: A text string to describe the nature of the list
                        elements. For example, 'lecture'


    :return: a key that is mapped with the templates/lectures dictionaries
    :rtype: string

    :raises Assertion ValueError: if the input when choosing the element
                                  is not in the range of the list
                                  is not a number
                                  is a negative number
    """
    for i in range(len(mapping_num_to_key_dict)):
        print(str(i) + " = [" + mapping_num_to_key_dict[i] + "]")

    num = general_get_input(
        "\nEnter the number of the " + item + " to be used: ")
    try:
        num = int(num)
        if num >= len(mapping_num_to_key_dict) or num < 0:
            raise ValueError()
    except ValueError:
        print(
            MapPre.colored("\nERROR: ", "red")
            + "Please input an integer between the range of 0 - "
            + str(len(mapping_num_to_key_dict) - 1)
        )
        general_get_input("Please press Any key to close the program\n")
        sys.exit(1)
    return mapping_num_to_key_dict[int(num)]


if __name__ == "__main__":
    tag_for_new_lect = "newVersion"
    tag_for_templates = "template"

    # ask whether the user would like to use the directory where the program
    # is located to find lectures,
    # templates and store the new lectures
    print("This is the Refactor tool. \nSee README.md for the functionality.")
    print("Make sure no PPT files are open.")
    print("Make sure OneDrive and any oter backup apps are paused.")
    ans = general_get_input(
        "\nPress [y] if you would like to use "
        + "the directory where this program is located"
        + " to find lectures, templates and store"
        + " the new presentations\n"
        + "Notice that this option will recursively"
        + " look for templates that should contain the word ["
        + tag_for_templates
        + "] in their file name.\n"
        + "Otherwise, please press any key to continue"
        + " and input custom directories\n"
        + "\n\nEnter a character:"
    ).lower()
    # Recollect lectures and templates before mapping
    if ans == "y":
        default_dir = os.path.abspath(os.getcwd())
        dict_lectures, dict_templates = find_powerpoint_lectures_and_templates(
            default_dir, tag_for_templates
        )
        dir_new_lect = default_dir

    else:
        dir_lectures = user_input_path(
            "Enter a folder path to extract lectures: ")
        dict_lectures = find_powerpoint_files(dir_lectures)

        dir_templates = user_input_path(
            "Enter a folder path to extract templates: ")
        dict_templates = find_powerpoint_files(dir_templates)
        dir_new_lect = user_input_path(
            "Enter a folder path to save the new lectures: ")

    # Sorting lectures presentations in natural order
    # for a neat display in the terminal or cmd
    mapping_number_to_key_dict_template = natsort.natsorted(dict_templates)
    mapping_number_to_key_dict_lectures = natsort.natsorted(dict_lectures)

    dir_new_lect = os.path.join(dir_new_lect, "newSlides")
    try:
        if os.path.isdir(dir_new_lect):
            raise ValueError()
    except ValueError:
        print(
            MapPre.colored("\nERROR: ", "red")
            + "\nThere is an existing newSlides folder which clashes with "
            + "the default output folder name."
            + "\nPlease rename your folder to something else."
            + "\nQuitting program now."
        )
        sys.exit(1)

    make_new_dir(dir_new_lect)

    template_key = select_a_number(
        mapping_number_to_key_dict_template, "template")

    is_changing_template = False
    max_lect_to_map = 0
    min_lect_to_map = 0
    lecture_key = ""
    name_modification = ""
    new_template_selected = ""
    old_lecture_selected = ""
    pdf_option = 0

    # the program will run in a while loop, letting the user choose between
    # generating only .pptx or .pptx and .pdf files, then to choose between
    # map all the lectures found in the directory, changing its template
    # or mapping an specific lecture
    while True:

        # before mapping starts, the user is asked to pick saving option
        # pptx only input 0; pptx and pdf input 1
        pdf_gen_on = general_get_input(
            "[q] for exit"
            + "\n[p] for generating .pptx files only"
            + "\n[b] for generating .pptx and .pdf files"
            + "\nAny other key for default - .pptx only"
            + "\n\nEnter your selection: "
        ).lower()
        if pdf_gen_on == "q":
            break
        if pdf_gen_on == "p":
            print("Mapping only to.pptx format")
            pdf_option = 0
        if pdf_gen_on == "b":
            print("Mapping to both formats")
            pdf_option = 1

        # choosing between  map all the lectures found in the directory,
        # changing its template or mapping an specific lecture
        letter = general_get_input(
            "[q] for exit"
            + "\n[a] for mapping all presentations in the current directory"
            + "\n[c] for changing the template \nAny other key to select"
            + " a specific presentation\n\nEnter your selection: "
        ).lower()
        if letter == "q":
            break
        elif letter == "a":
            min_lect_to_map = 0
            max_lect_to_map = len(mapping_number_to_key_dict_lectures)
        elif letter == "c":
            is_changing_template = True
        else:
            min_lect_to_map = 0
            max_lect_to_map = 1

        if is_changing_template is True:
            print("current template is: ", template_key)
            template_key = select_a_number(
                mapping_number_to_key_dict_template, "template"
            )
            is_changing_template = False
            print("The template has been changed to: ", template_key)
        else:
            # in the case that the user decides to map all lectures
            # max_lect_to_map will be set up the maximum number
            # of mapping_number_to_key_dict_lectures
            # whereas min_lect_to_map to zero
            # so that we can use a while to extract the path
            # from the lectures and extra information
            # needed for mappingPreV2 method from MapPre
            # Notice if there is only one lecture contained in the directory
            # this will be automatically mapped when selecting a specific
            # presentation as an option
            while min_lect_to_map < max_lect_to_map:
                print("\nProcessing...\nDO NOT TOUCH THE PPT FILES.\n")
                if max_lect_to_map == len(mapping_number_to_key_dict_lectures):
                    lecture_key = mapping_number_to_key_dict_lectures[
                        min_lect_to_map]
                else:
                    lecture_key = select_a_number(
                        mapping_number_to_key_dict_lectures, "lecture"
                    )
                new_template_selected = dict_templates[template_key]
                old_lecture_selected = dict_lectures[lecture_key]

                print(
                    "\nProcessing lecture "
                    + str(lecture_key)
                    + " with template "
                    + str(template_key)
                    + " now..."
                )

                # create subfolder name for output lecture PPT
                new_lect_subfolder = os.path.dirname(lecture_key)
                new_lect_filename = os.path.basename(lecture_key)
                make_new_dir(os.path.join(dir_new_lect, new_lect_subfolder))
                name_modification = (
                    dir_new_lect
                    + "/"
                    + new_lect_subfolder
                    + "/"
                    + tag_for_new_lect
                    + "-"
                    + new_lect_filename
                )

                print("Generating output for " + str(name_modification) + "\n")

                MapPre.mappingPreV2(
                    old_lecture_selected,
                    new_template_selected,
                    name_modification,
                    dir_new_lect,
                    lecture_key,
                    template_key,
                    pdf_option,
                )
                min_lect_to_map += 1

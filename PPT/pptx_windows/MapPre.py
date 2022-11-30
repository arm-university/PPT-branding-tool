"""
This confidential and proprietary software may be used only as
authorised by a licensing agreement from ARM Limited
(C) COPYRIGHT 2020 ARM Limited
ALL RIGHTS RESERVED
The entire notice above must be reproduced on all authorised
copies and copies may only be made to the extent permitted
by a licensing agreement from ARM Limited.
Author: Education Team
Date: 28/07/2020
Summary: This piece of code contains the initilisation of termcolors and
         a function that applies a new template to an old presentation
"""
from colorama import init
from termcolor import colored
import win32com.client
import win32gui
import win32con
import time

# initilisation of termcolor to work on Windows systems
init()


def mappingPreV2(
    path_lecture,
    path_template,
    name_for_new_lect,
    path_for_new_lecture,
    name_for_lect,
    name_for_temp,
    pdf_on,
):
    """ this function will use the win32 library to apply a new template
        PowerPoint to a lecture presentation

        :param string path_lecture: location of the old lecture
        :param string path_template: location of the desired template
        :param string name_for_new_lect: name that will used the new lecture
        :param string path_for_new_lecture: location to store the new lecture
        :param string name_for_lect: name of the old lecture
        :param string name_for_temp: name of the template
        :param int pdf_on: binary variable to enable .pdf file generation

        :return: void

         PLEASE CLOSE ANY PowerPoint processes in your machine before running
         this function
    """
    background_slide_warnings = []
    dest_path = path_template
    source_path = path_lecture

    # create an instance of PowerPoint application by default
    # the instance will be visible so we make invisible using win32gui
    ppt_instance = win32com.client.Dispatch("PowerPoint.Application")
    ppt_instance.Visible = True
    hwnd = win32gui.FindWindow(None, "PowerPoint")
    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

    # we use the instance created above to open two different presentations
    # or the lecture and template respectively, the instance will be visible
    # so use the Open method to set the visibility of this window to false
    prs = ppt_instance.Presentations.Open(source_path, WithWindow=False)
    prs2 = ppt_instance.Presentations.Open(dest_path, WithWindow=False)

    slide_count = len(prs.Slides)
    prs2.Slides.InsertFromFile(source_path, 0, 1, slide_count)

    # sometimes copying the content from one presentation to another one
    # will result in skipping non-master background slides.
    # to avoid this:
    # -Go through all slides
    # -Check that all of them have a background from slide master
    # -If not, take all the shapes and content from that slide
    # -Make a capture, store and reload it in the new presentation
    for i in range(slide_count):
        if prs.Slides[i].FollowMasterBackground == 0:
            background_slide_warnings.append(str(i + 1))
            prs2.Slides[i].FollowMasterBackground = 0
            for shape in prs.Slides[i].Shapes:
                shape.Visible = 0
            prs.Slides[i].DisplayMasterShapes = 0
            # prs.Slides[i].Export function will need two arguments for the
            # width and height of the photo, as long as we
            # use large numbers for these arguments
            # we will obtain a good quality photo of the background slide
            prs.Slides[i].Export(
                path_for_new_lecture + "/slide.png", "PNG", 13333, 7500
            )
            prs2.Slides[i].FollowMasterBackground = 0
            prs2.Slides[i].Background.Fill.UserPicture(
                path_for_new_lecture + "/slide.png"
            )
    # save the lecture as pdf (if enabled) and PowerPoint documents
    if pdf_on == 1:
        prs2.SaveAs(name_for_new_lect[0:len(name_for_new_lect) - 5]
                    + ".pdf", 32)
    prs2.SaveAs(name_for_new_lect)

    # close everything
    prs.Close()
    prs2.Close()
    # kill all instances of PowerPoint
    ppt_instance.Quit()
    del ppt_instance

    # output any warnings or errors here
    if len(background_slide_warnings) > 0:
        print(
            colored(" WARNING: ", "yellow")
            + "Presentation:"
            + name_for_lect
            + ":\nThe following slides contained background images: ",
            end=" ",
        )
        for i in background_slide_warnings:
            print(i, end=" ")
    print(
        colored(
            "\n**************************************************************",
            "green",
        )
    )

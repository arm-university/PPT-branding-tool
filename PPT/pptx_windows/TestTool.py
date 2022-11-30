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
Summary: This script contain unit tests for:
         -Refactor.py
         -MapPre.py
"""
import Refactor
import filecmp
import os
import unittest
from unittest import TestCase
from unittest.mock import patch, call, Mock


class TestRefactor(TestCase):
    def setUp(self):
        """ this function will set up the environment before running a test

            :param self self: instance of the class

            :return: void
        """
        # not sure whether to have this as a class variable
        # let me know in the review
        self.mock_mapping_list = ["a", "b", "c"]

    @patch("Refactor.os.walk")
    def test_find_powerpoint_files_in_a_folder_with_data(self, mock_walk):
        """ this function will test the method find_powerpoint_files
            from Refactor, it will be tested with a mock folder structure
            the test will compare whether the constructed dictionary
            is identical to the expected one
            os.walk will use the mock mock_tuple_walk
            as a tree directory and will ONLY pick the pptx files

            :param self self: instance of the class
            :param mock mock_walk: it mocks Refact.os.walk method

            :return: void

            :raises Assertion Error: if find_powerpoint_files method does not
                                     return the expected dictionary
        """
        # mock_tuple_walk dir will look like this
        # |mock
        # --|module1
        # ----|lecture1.pptx
        # ----|garbage_that_should_not_be_picked.txt
        # ----|~$hidden_temporary_file_that_should_not_be_picked.pptx
        # ----|lecture2.pptx
        # --|module2
        # ----|lecture2.pptx
        # --|module3
        # ----|lecture3.pptx
        # ----|garbage_that_should_not_be_picked.docx
        # ----|~$hidden_temporary_file_that_should_not_be_picked.pptx
        # -------|module3a
        # ----------|lecture1.pptx

        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3"), ("module3a")),
            (
                "/mock/module1",
                (),
                (
                    "lecture1.pptx",
                    "garbage_that_should_not_be_picked.txt",
                    "~$hidden_temporary_file_that_should_not_be_picked.pptx",
                    "lecture2.pptx"
                ),
            ),
            ("/mock/module2", (), ("lecture2.pptx",)),
            (
                "/mock/module3",
                (),
                (
                    "lecture3.pptx",
                    "garbage_that_should_not_be_picked.docx",
                    "~$hidden_temporary_file_that_should_not_be_picked.pptx",
                ),
            ),
            (
                "/mock/module3/module3a",
                (),
                (
                    "lecture1.pptx",
                ),
            ),
        ]
        for aTuple in mock_tuple_walk:
            print(aTuple)
        dict_expected = {
            os.path.join("module1", "lecture1.pptx"): os.path.join(
                "/mock/module1", "lecture1.pptx"),
            os.path.join("module1", "lecture2.pptx"): os.path.join(
                "/mock/module1", "lecture2.pptx"),
            os.path.join("module2", "lecture2.pptx"): os.path.join(
                "/mock/module2", "lecture2.pptx"),
            os.path.join("module3", "lecture3.pptx"): os.path.join(
                "/mock/module3", "lecture3.pptx"),
            os.path.join("module3", "module3a", "lecture1.pptx"): os.path.join(
                "/mock/module3/module3a", "lecture1.pptx"),
        }
        mock_walk.return_value = mock_tuple_walk
        self.assertDictEqual(
            Refactor.find_powerpoint_files("/mock"),
            dict_expected,
        )

    @patch("Refactor.os.walk")
    def test_find_powerpoint_files_in_a_folder_with_no_data(self, mock_walk):
        """ this function will test the method find_powerpoint_files
            from Refactor to check for the raising of an exception
            when this method is applied to a folder with no data,
            os.walk will use the mock_tuple_walk
            as a tree ditectory and will ONLY pick the pptx files

            :param self self: instance of the class
            :param mock mock_walk: it mocks Refact.os.walk method

            :return: void

            :raises Assertion Error: if find_powerpoint_files method does not
                                     raise an exception due to not finding
                                     pptx files in the current dir
        """
        # mock_tuple_walk dir will look like this
        # |mock
        # --|module1
        # --|module2
        # --|module3
        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3"), ()),
        ]
        mock_walk.return_value = mock_tuple_walk
        with self.assertRaises(SystemExit) as cm:
            Refactor.find_powerpoint_files("user_input_mock_path")
            self.assertEqual(cm.exception, "Error")

    @patch("Refactor.os.walk")
    def test_find_powerpoint_lectures_and_templates_with_data(self, mock_walk):
        """ this function will test the method
            find_powerpoint_lectures_and_templates from Refactor
            we assume that the directory that will be transverse
            has write/read permission
            and that templates presentation names contain a tag
            so that they can be distinguished between lectures
            os.walk will use the mock_tuple_walk
            as a tree ditectory and will ONLY pick the pptx files

            :param self self: instance of the class
            :param mock mock_walk: it mocks Refact.os.walk method

            :return: void

            :raises Assertion Error: if find_powerpoint_files method does not
                                     return the expected dictionary
        """
        # mock_tuple_walk dir will look like this
        # |mock
        # --|module1
        # ----|lecture1.pptx
        # ----|garbage_that_should_not_be_picked.txt
        # ----|~$hidden_temporary_file_that_should_not_be_picked.pptx
        # ----|lecture2.pptx
        # --|module2
        # ----|lecture2.pptx
        # --|module3
        # ----|lecture3.pptx
        # ----|garbage_that_should_not_be_picked.docx
        # -------|module3a
        # ----------|lecture1.pptx
        # --|template
        # ----|templatetag-arm2020.pptx
        # ----|garbage_that_should_not_be_picked.docx
        # -----|~$hidden_temporary_file_that_should_not_be_picked.pptx
        # ----|templatetag-arm2021.pptx.docx
        # -------|template_subfolder
        # ----------|templatetag-arm2020.pptx
        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3", "template"), (
                "module3a", "template_subfolder")),
            (
                "/mock/module1",
                (),
                (
                    "lecture1.pptx",
                    "garbage_that_should_not_be_picked.txt",
                    "~$hidden_temporary_file_that_should_not_be_picked.pptx",
                    "lecture2.pptx"
                ),
            ),
            ("/mock/module2", (), ("lecture2.pptx",)),
            (
                "/mock/module3",
                (),
                ("lecture3.pptx", "garbage_that_should_not_be_picked.docx"),
            ),
            (
                "/mock/module3/module3a",
                (),
                (
                    "lecture1.pptx",
                ),
            ),
            (
                "/mock/template",
                (),
                (
                    "templatetag-arm2020.pptx",
                    "garbage_that_should_not_be_picked.docx",
                    "~$hidden_temporary_file_that_should_not_be_picked.pptx",
                    "templatetag-arm2021.pptx",
                ),
            ),
            (
                "/mock/template/template_subfolder",
                (),
                (
                    "templatetag-arm2020.pptx",
                ),
            ),
        ]
        for aTuple in mock_tuple_walk:
            print(aTuple)
        dict_lect_expected = {
            os.path.join("module1", "lecture1.pptx"): os.path.join(
                "/mock/module1", "lecture1.pptx"),
            os.path.join("module1", "lecture2.pptx"): os.path.join(
                "/mock/module1", "lecture2.pptx"),
            os.path.join("module2", "lecture2.pptx"): os.path.join(
                "/mock/module2", "lecture2.pptx"),
            os.path.join("module3", "lecture3.pptx"): os.path.join(
                "/mock/module3", "lecture3.pptx"),
            os.path.join("module3", "module3a", "lecture1.pptx"): os.path.join(
                "/mock/module3/module3a", "lecture1.pptx"),
        }
        dict_template_expected = {
            os.path.join("template", "templatetag-arm2020.pptx"): os.path.join(
                "/mock/template", "templatetag-arm2020.pptx"
            ),
            os.path.join("template", "templatetag-arm2021.pptx"): os.path.join(
                "/mock/template", "templatetag-arm2021.pptx"
            ),
            os.path.join(
                "template", "template_subfolder", "templatetag-arm2020.pptx"
                ): os.path.join(
                    "/mock/template/template_subfolder",
                    "templatetag-arm2020.pptx"),
        }
        mock_walk.return_value = mock_tuple_walk
        (
            result_dict_lect,
            result_dict_template,
        ) = Refactor.find_powerpoint_lectures_and_templates(
            "/mock", "templatetag"
        )
        self.assertDictEqual(result_dict_lect, dict_lect_expected)
        self.assertDictEqual(result_dict_template, dict_template_expected)

    @patch("Refactor.os.walk")
    def test_find_powerpoint_lectures_and_templates_no_data(self, mock_walk):
        """ this function will test the method
            find_powerpoint_lectures_and_templates from Refactor
            to check for the raising of an exception
            when this method is applied to a folder with no data,
            we assume that the directory that will be transverse
            has write/read permission
            and that templates presentation names contain a tag
            so that they can be distinguished between lectures
            os.walk will use the mock_tuple_walk
            as a tree ditectory and will ONLY pick the pptx files

            :param self self: instance of the class
            :param mock mock_walk: it mocks Refact.os.walk method

            :return: void

            :raises Assertion Error: if find_powerpoint_lectures_and_templates
                                     method does not raise an exception
                                     due to not finding pptx lectures
                                     or templates while
                                     transversing the directory
        """
        # mock_tuple_walk dir will look like this
        # |mock
        # --|module1
        # ----|lecture1.pptx
        # --|module2
        # ----|lecture2.pptx
        # --|module3
        # ----|lecture3.pptx
        # --|template
        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3", "template"), ()),
            ("/mock/module1", (), ("lecture1.pptx",),),
            ("/mock/module2", (), ("lecture2.pptx",),),
            ("/mock/module3", (), ("lecture3.pptx",),),
        ]
        mock_walk.return_value = mock_tuple_walk
        path = "user_input_mock_path"

        # check method raises exception
        # because directory doesn't contain templates
        with self.assertRaises(SystemExit) as cm:
            Refactor.find_powerpoint_lectures_and_templates(path, "template")
            self.assertEqual(cm.exception, "Error")

        # mock_tuple_walk dir will look like this
        # |mock
        # --|template
        # ----|template-arm2020.pptx
        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3", "template"), ()),
            ("/mock/template", (), ("template-arm2020.pptx",),),
        ]
        mock_walk.return_value = mock_tuple_walk

        # check method raises exception
        # because directory doesn't contain lectures
        with self.assertRaises(SystemExit) as cm:
            Refactor.find_powerpoint_lectures_and_templates(path, "template")
            self.assertEqual(cm.exception, "Error")

        # mock_tuple_walk dir will look like this
        # |mock
        # --|module1
        # --|module2
        # --|module3
        # --|template
        mock_tuple_walk = [
            ("/mock", ("module1", "module2", "module3", "template"), ()),
        ]
        mock_walk.return_value = mock_tuple_walk

        # check method raises exception because directory doesn't
        # contain lectures and templates
        with self.assertRaises(SystemExit) as cm:
            Refactor.find_powerpoint_lectures_and_templates(path, "template")
            self.assertEqual(cm.exception, "Error")

    @patch("Refactor.os.path")
    @patch("Refactor.os")
    def test_make_new_dir_when_it_does_not_exists(self, mock_os, mock_path):
        """ this function will test the method make_new_dir
            from Refactor when a directory that the user wants
            to create does not exists

            :param self self: instance of the class
            :param mock mock_os: it mock Refactor.os method
            :param mock mock_path: it mocks Refactor.os.path method

            :return: void

            :raises Assertion Error: if there is not an assertion of makedirs
                                     creating a directory
        """
        mock_path.exists.return_value = False
        dir_path = os.path.join("user_input_mock_path", "newSlides")
        Refactor.make_new_dir(dir_path)
        mock_os.makedirs.assert_called_with(dir_path)

    @patch("Refactor.os.path")
    @patch("Refactor.os")
    def test_make_new_dir_when_it_exists(self, mock_os, mock_path):
        """ this function will test the method make_new_dir
            from Refactor when a directory that the user wants
            to create exists


            :param self self: instance of the class
            :param mock mock_os: it mock Refactor.os method
            :param mock mock_path: it mocks Refactor.os.path method

            :return: void

            :raises Assertion Error: if there is 1 or more calls of makedirs
                                     coming out from make_new_dir method
        """
        mock_path.exists.return_value = True
        dir_path = "user_input_mock_path"
        Refactor.make_new_dir(dir_path)
        assert mock_os.makedirs.call_count == 0

    @patch("Refactor.MapPre.win32com.client.Dispatch")
    def test_mapping_method_without_background_images_pdf_on_off(
        self, mock_win32com_client_dispatch
    ):
        """ this function will test the method mappingPreV2
            imported from MapPre.py into Refactor
            we are assuming that the library win32 has already been tested
            and the only thing we need to do is to check the calls done
            in this method, we assume the pptx does not contain
            background images


            :param self self: instance of the class
            :param mock mock_win32com_client_dispatch: it mocks
                                Refactor.MapPre.win32com.client.dispatch method

            :return: void

            :raises Assertion Error: if one of the calls was not done
                                     when calling mappingPreV2 method
        """
        # To avoid using real ppptx files, we are mocking the calls return
        # of a presentation that will contain 1 slide and will
        # use a FollowMasterBackground
        print("Now we test pdf on or off without background images")
        for x in range(2):
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides.__len__.return_value = 1
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides[0].FollowMasterBackground = 1
            Refactor.MapPre.mappingPreV2(
                "mock_lect",
                "mock_temp",
                "mock_new_file.pptx",
                "mock_new_path",
                "mock_lect_name.pptx",
                "mock_temp_name.pptx",
                x,
            )
            # we split the calls into three different lists
            # because mock_win32com_client_dispatch does not allow skip
            # calls_part2 unless specified, and in this case
            # calls_part1 and calls_part3 could be checked in order
            calls_part1 = [
                call("PowerPoint.Application"),
                call().Presentations.Open("mock_lect", WithWindow=False),
                call().Presentations.Open("mock_temp", WithWindow=False),
            ]

            mock_win32com_client_dispatch.assert_has_calls(calls_part1)
            # pdf option on
            if x == 1:
                calls_part2 = [
                    call()
                    .Presentations.Open("mock_temp")
                    .Slides.InsertFromFile("mock_lect", 0, 1, 1),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pdf", 32),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pptx"),
                ]
                mock_win32com_client_dispatch.assert_has_calls(
                    calls_part2, any_order=True
                )
            # pdf option off
            elif x == 0:
                calls_part2 = [
                    call()
                    .Presentations.Open("mock_temp")
                    .Slides.InsertFromFile("mock_lect", 0, 1, 1),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pptx"),
                ]
                mock_win32com_client_dispatch.assert_has_calls(
                    calls_part2, any_order=True
                )

            calls_part3 = [
                call().Presentations.Open("mock_lect").Close(),
                call().Presentations.Open("mock_temp").Close(),
                call().Quit(),
            ]
            mock_win32com_client_dispatch.assert_has_calls(calls_part3)

    @patch("Refactor.MapPre.win32com.client.Dispatch")
    def test_mapping_method_with_background_images_pdf_on_off(
        self, mock_win32com_client_dispatch
    ):
        """ this function will test the method mappingPreV2
            imported from MapPre.py into Refactor
            we are assuming that the library win32 has already been tested
            and the only thing we need to do is to check the calls done
            in this method, we assume the pptx contains
            background images


            :param self self: instance of the class
            :param mock mock_win32com_client_dispatch: it mocks
                                Refactor.MapPre.win32com.client.dispatch method
            :return: void

            :raises Assertion Error: if one of the calls was not done
                                     when calling mappingPreV2 method
        """
        print("Now we test pdf on or off with background images")
        for x in range(2):
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides.__len__.return_value = 1
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides[0].FollowMasterBackground = 0
            Refactor.MapPre.mappingPreV2(
                "mock_lect",
                "mock_temp",
                "mock_new_file.pptx",
                "mock_new_path",
                "mock_lect_name.pptx",
                "mock_temp_name.pptx",
                x,
            )
            # we split the calls into three different lists
            # because mock_win32com_client_dispatch does not allow skip
            # calls_part2 unless specified, and in this case
            # calls_part1 and calls_part3 could be checked in order
            calls_part1 = [
                call("PowerPoint.Application"),
                call().Presentations.Open("mock_lect", WithWindow=False),
                call().Presentations.Open("mock_temp", WithWindow=False),
            ]
            mock_win32com_client_dispatch.assert_has_calls(calls_part1)
            # pdf option on
            if x == 1:
                calls_part2 = [
                    call()
                    .Presentations.Open("mock_temp")
                    .Slides.InsertFromFile("mock_lect", 0, 1, 1),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pdf", 32),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pptx"),
                ]
                mock_win32com_client_dispatch.assert_has_calls(
                    calls_part2, any_order=True
                )
            # pdf option off
            elif x == 0:
                calls_part2 = [
                    call()
                    .Presentations.Open("mock_temp")
                    .Slides.InsertFromFile("mock_lect", 0, 1, 1),
                    call()
                    .Presentations.Open("mock_temp")
                    .SaveAs("mock_new_file.pptx"),
                ]
                mock_win32com_client_dispatch.assert_has_calls(
                    calls_part2, any_order=True
                )

            calls_part3 = [
                call().Presentations.Open("mock_lect").Close(),
                call().Presentations.Open("mock_temp").Close(),
                call().Quit(),
            ]
            mock_win32com_client_dispatch.assert_has_calls(calls_part3)

            # some extra calls that should be checked when the pptx contains
            # background images
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides[0].Export.assert_called_with(
                "mock_new_path/slide.png", "PNG", 13333, 7500
            )
            mock_win32com_client_dispatch(
                "PowerPoint.Application"
            ).Presentations.Open().Slides[
                0
            ].Background.Fill.UserPicture.assert_called_with(
                "mock_new_path/slide.png"
            )
    # in order to mock an input insertation from the user, we will use
    # Refact.general_get_input where
    # we set up the return_value to be any string

    @patch("Refactor.general_get_input", return_value="mock_user_input")
    @patch("Refactor.os.path")
    def test_user_input_path_success(self, mock_path, input):
        """ this function will test the method user_input_path from Refactor
            when an existent path is provided

            :param self self: instance of the class
            :param mock mock_path: it mocks Refactor.os.path method
            :param mock input: it mocks the Refactor.general_get_input method

            :return: void

            :raises Assertion Error:-if user_input_path method does not
                                     return expected string
                                    -if mock_path is not called with expected
                                     string
        """

        # set the call  Refactor.os.path.exists to true so that
        # we can test user_input_path return
        mock_path.exists.return_value = True
        # in user_input_path, there is a call to general_get_input which is
        # is waiting for user input, because we are mocking this input,
        # the input being passed is equal to the return value
        # defined in the @path above the function
        self.assertEqual(Refactor.user_input_path(""), "mock_user_input")
        mock_path.exists.assert_called_with("mock_user_input")

    @patch("Refactor.general_get_input", return_value="mock_user_input")
    @patch("Refactor.os.path")
    def test_user_input_path_error(self, mock_path, input):
        """ this function will test the method user_input_path from Refactor
            when a non-existent path is provided

            :param self self: instance of the class
            :param mock mock_path: it mocks Refactor.os.path method
            :param mock input: it mocks the Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if user_input_path method does not
                                     raise a Exception
        """
        # set the call of exists to false so that
        # we can have an exception coming out from the user_input_path method
        mock_path.exists.return_value = False
        self.assertRaises(Exception, lambda: Refactor.user_input_path(""))

    @patch("Refactor.general_get_input", return_value="0")
    def test_select_a_number_first_element(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send the selection of the element

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if the expected element is not the one
                                     returned by the select_a_number method
        """
        self.assertEqual(
            Refactor.select_a_number(self.mock_mapping_list, "lect"),
            self.mock_mapping_list[0],
        )

    @patch("Refactor.general_get_input", return_value="1")
    def test_select_a_number_middle_element(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send the selection of the element

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if the expected element is not the one
                                     returned by the select_a_number method
        """
        self.assertEqual(
            Refactor.select_a_number(self.mock_mapping_list, "lect"),
            self.mock_mapping_list[1],
        )

    @patch("Refactor.general_get_input", return_value="2")
    def test_select_a_number_last_element(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send the selection of the element

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if the expected element is not the one
                                     returned by the select_a_number method
        """
        self.assertEqual(
            Refactor.select_a_number(self.mock_mapping_list, "lect"),
            self.mock_mapping_list[2],
        )

    @patch("Refactor.general_get_input", return_value="3")
    def test_select_a_number_overflow_element(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send the a number greater than the list
            length which should raise an exception

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if there is not an exception when sending
                                     coming out from select_a_number
        """
        with self.assertRaises(SystemExit) as cm:
            Refactor.select_a_number(self.mock_mapping_list, "lect")
            self.assertEqual(cm.exception, "Error")

    @patch("Refactor.general_get_input", return_value="-1")
    def test_select_a_number_negative(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send a negative number
            which should raise an exception

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if there is not an exception when sending
                                     coming out from select_a_number
        """
        with self.assertRaises(SystemExit) as cm:
            Refactor.select_a_number(self.mock_mapping_list, "lect")
            self.assertEqual(cm.exception, "Error")

    @patch("Refactor.general_get_input", return_value="word")
    def test_select_a_number_no_number(self, input):
        """ this function will test the method select_a_number
            from Refactor, it will use the defined mock_mapping_list as
            one the list of elements and use the wrapper method
            general_get_input to send the a non number which should
            raise an exception

            :param self self: instance of the class
            :param mock input: it mocks Refactor.general_get_input method

            :return: void

            :raises Assertion Error: if there is not an exception when sending
                                      coming out from select_a_number
        """
        with self.assertRaises(SystemExit) as cm:
            Refactor.select_a_number(self.mock_mapping_list, "lect")
            self.assertEqual(cm.exception, "Error")


if __name__ == "__main__":
    unittest.main()

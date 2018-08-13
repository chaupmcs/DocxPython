import json

from gevent import os
from docx_python import DOCXPython
import time

def test_many_templates(project_path, templates):
    files_in_rels = os.listdir(templates)


    for file_name in files_in_rels:
        a = DOCXPython(project_path, file_name)
        a.get_new_word_template(file_name, False)

        a.replace_header_and_footer()
        a.replace_signature()
        a.replace_mailing_merge_and_convert_to_pdf()
        a.print_finish()

    print("=======================================================")


def test_one_template(project_path, file_name, convert_to_pdf, multiple_files):


    DOCXPython.clear_all_data(project_path)

    a = DOCXPython(project_path, file_name, convert_to_pdf, multiple_files)
    a.get_new_word_template(file_name, False)

    a.replace_header_and_footer()
    a.replace_signature()
    a.replace_mailing_merge_and_convert_to_pdf()
    a.print_finish()

    #DOCXPython.clear_all_data(project_path)
    print("=======================================================")


def len_dataMailingMerge(project_path):
    path = os.path.join(project_path, 'input_data', 'data_mailingMerge.json')
    with open(path) as f:
        data = json.load(f)['data']
    return len(data)

if __name__ == '__main__':

    #########################   Edit parameters here  #########################
    project_path = '/Users/minhchau/Downloads/DocxPython'
    file_name = 'word_template.docx'
    convert_to_pdf = False
    multiple_files = True




    ##########################################################################################
    DOCXPython.clear_all_data(project_path)
    starting_time = time.time()
    test_one_template(project_path, file_name, convert_to_pdf, multiple_files)
    print("--- Done.  Total run time: %s seconds ---" % (time.time() - starting_time))
    print("len of data_mailing_merge is {}".format(len_dataMailingMerge(project_path)))
    #DOCXPython.clear_all_data(project_path)
    ##########################################################################################

    # --- Done.Total
    # run 65 files
    # convert_to_pdf: time: 581.8146390914917 seconds
    # do not convert_to_pdf: time 1.786578893661499 seconds


    # 325 files:
    # do not convert_to_pdf: multiple files: 7.670944929122925 seconds, 1 files: 2.42118501663208

    # --- Done.Total
    # run
    # time: 36.15711808204651
    # seconds - --
    # len
    # of
    # data_mailing_merge is 1625
    # multiple file, dont convert

    # -- get a new word template --
    # replace_header_and_footer is done
    # replace_signature is done !
    # creating one file contaning 1625 mails...
    # replace_mailingMerge_and_convertToPDF is done !
    # == == == == == == == == == == == == == == == == == == == == == == == == == == == =
    # --- Done.Total
    # run
    # time: 11.138234853744507
    # seconds - --
    # ===========================



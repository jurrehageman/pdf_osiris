import parse_pdf
import argparse
import os

def get_comm_args():
    """
    Reads command line arguments
    """
    parser = argparse.ArgumentParser(
        description="Batch parse pdf Osiris")     
    parser.add_argument("pdf_folder",
                        help="the path to the pdf folder")
    parser.add_argument("excel_folder",
                        help="the path to the excel folder")                    
    parser.add_argument('-v', '--verbose', action='store_true',
                        help="verbose mode")
    args = parser.parse_args()
    return args


def check_folder_exists(folder):
    return os.path.exists(folder)


def create_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def read_folder(verbose_status, in_folder, out_folder, slb_list):
    for filename in os.listdir(in_folder):
        if filename.endswith(".pdf"):
            head, tail = os.path.splitext(filename)
            file_path = os.path.join(in_folder, filename)
            pdf_content = parse_pdf.extract_pdf(file_path)
            student_data = parse_pdf.parse_text(pdf_content)
            create_dir(out_folder)
            out_file_path = os.path.join(out_folder, head)
            print("now working on:", out_file_path)
            student_stats = parse_pdf.get_student_stats(student_data, slb_list)
            if verbose_status:
                parse_pdf.print_stats(student_data, student_stats, slb_list)
            parse_pdf.write_excel(student_data, student_stats, slb_list, out_file_path)
            

def main():
    args = get_comm_args()
    verbose_status = args.verbose
    in_folder = args.pdf_folder
    out_folder = args.excel_folder
    slb_list = parse_pdf.read_slb(parse_pdf.slb_file)
    if not check_folder_exists(in_folder):
        print(in_folder, "not found")
        return
    read_folder(verbose_status, in_folder, out_folder, slb_list)
    print("Data written to Excel folder")
    print("Done")
    print("*" * 20)

if __name__ == "__main__":
    main()
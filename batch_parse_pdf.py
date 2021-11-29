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
    args = parser.parse_args()
    return args


def check_folder_exists(folder):
    return os.path.exists(folder)


def create_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def read_folder(folder, slb_list):
    for filename in os.listdir(folder):
        if filename.endswith(".pdf"):
            head, tail = os.path.splitext(filename)
            file_path = os.path.join(folder, filename)
            pdf_content = parse_pdf.extract_pdf(file_path)
            student_data = parse_pdf.parse_text(slb_list, pdf_content)
            create_dir("excel")
            out_file_path = os.path.join("excel", head)
            print("now working on:", out_file_path)
            parse_pdf.write_excel(student_data, out_file_path)
            

def main():
    args = get_comm_args()
    in_folder = args.pdf_folder
    slb_list = parse_pdf.read_slb(parse_pdf.slb_file)
    if not check_folder_exists(in_folder):
        print(in_folder, "not found")
        return
    read_folder(in_folder, slb_list)
    print("Data written to Excel folder")
    print("Done")
    print("*" * 20)

if __name__ == "__main__":
    main()
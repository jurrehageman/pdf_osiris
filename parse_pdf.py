import pdfplumber
import argparse
import xlsxwriter

# Globals
slb_file = "slb.csv" 

def get_comm_args():
    """
    Reads command line arguments
    """
    parser = argparse.ArgumentParser(
        description="Parse pdf Osiris")     
    parser.add_argument("pdf_file",
                        help="the path to the pdf File downloaded from Osiris")
    parser.add_argument("output_file",
                        help="the path to the csv file with the output")
    parser.add_argument('-v', '--verbose', action='store_true',
                        help="verbose mode")
    args = parser.parse_args()
    return args


def read_slb(slb_file):
    slb_list = []
    with open(slb_file) as f:
        for line in f:
            line = line.strip()
            if not line in slb_list:
                slb_list.append(line)
    return slb_list


def extract_pdf(infile):
    data = []
    with pdfplumber.open(infile) as pdf:
        for num, page in enumerate(pdf.pages):
            if num > 0:
                text = page.extract_text()
                data.append(text)
    return data


def parse_text(data):
    """This needs to be improved"""
    slb_dict = {}
    for line in data:
        student_list = []
        item = line.split("\n")
        item = [i for i in item if len(i) > 6] # get rid of BO (*) and CH (*) lines
        slb = item[2][-5: -1]
        students = item[6:]
        for student in students:
            student_dict = {}
            student_num = student[0:6].strip()
            comma_pos = student.find(",")
            last_name = student[7:comma_pos].strip()
            BO_pos = student.find(" BO ")
            first_name = student[comma_pos + 1: BO_pos].strip()
            if "(" in first_name:
                first_name = first_name[: first_name.find("(")] # sometimes no BO but comma
            first_name = first_name.split()
            first_name = ' '.join(i for i in first_name if not i.isupper()) # get rid of junk     
            student_dict['student_num'] = student_num
            student_dict['last_name'] = last_name
            student_dict['first_name'] = first_name
            student_dict['slb'] = slb
            student_list.append(student_dict)
        slb_dict[slb] = student_list
    return slb_dict


def print_stats(student_data, file_name, verbose_status, slb_list):
    print("currently parsing:", file_name)
    if verbose_status:
        unique_students = set()
        slb_ilst_students = set()
        slb_other = set()
        print()
        for slb in sorted(student_data):
            student_nums = [i["student_num"] for i in student_data[slb]]
            unique_students.update(student_nums)
            #print("{0:<6}{1:<3} students {2}".format(slb, len(student_data[slb]), student_nums))
            print("{0:<6}{1:<3} students".format(slb, len(student_data[slb])))
            if slb in slb_list:
                slb_ilst_students.update(student_nums)
            else:
                slb_other.update(student_nums)
        print()
        print("Total number of students:", len(unique_students))
        print("Total number of students, ILST SLBer:", len(slb_ilst_students), "(This will be written to Excel)")
        print("Total number of students, other SLBer:", len(slb_other))
        not_assigned = slb_other.difference(slb_ilst_students)
        print("Number of students other SLBer not assigned to ILST SLBer:", len(not_assigned))
        print()
        if not_assigned:
            print("Not assigned:")
            for slb in sorted(student_data):
                for student in student_data[slb]:
                    if student["student_num"] in not_assigned:
                        print(student)
    print()


def write_excel(slb_list, data, outfile):
    workbook = xlsxwriter.Workbook(outfile + ".xlsx")
    worksheet = workbook.add_worksheet()
    col = 0
    row = 1
    for slb in sorted(data):
        if slb in slb_list:
            for student in data[slb]:
                col = 0
                row_data = [student["student_num"], student["last_name"], student["first_name"], student["slb"]]
                for i in row_data:
                    worksheet.write(row, col, i)
                    col += 1
                row += 1
    worksheet.add_table(0, 0, row - 1, col - 1, {'columns': [{'header': 'Student Nummer'},
                                                            {'header': 'Achternaam'},
                                                            {'header': 'Voornaam'},
                                                            {'header': 'SLB-er'},
                                                            ]})
    workbook.close()


def main():
    args = get_comm_args()
    verbose_status = args.verbose
    in_file = args.pdf_file
    out_file = args.output_file
    slb_list = read_slb(slb_file)
    pdf_content = extract_pdf(in_file)
    student_data = parse_text(pdf_content)
    print_stats(student_data, in_file, verbose_status, slb_list)
    write_excel(slb_list, student_data, out_file)
    print()
    print("Data written to", out_file + ".xlsx")
    print("Done")
    print("*" * 20)


if __name__ == "__main__":
    main()
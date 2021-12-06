import pdfplumber
import argparse
import xlsxwriter
import collections

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
    student_data = {}
    for line in data:
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
            if not student_num in student_data:

                student_data[student_num] = {'last_name': last_name,
                                                'first_name': first_name,
                                                'slb': [slb]
                }
            else:
                student_data[student_num]['slb'].append(slb)
    return student_data


def print_stats(student_data, student_stats, slb_list):
    print("Stats:")
    print("Total number of students:", len(student_stats['all_students']))
    print("Total number of students, BML-R SLBer:", len(student_stats['bml_students']))
    print("Total number of students, other SLBer:", len(student_stats['other_students']))
    print("Number of students with multiple SLBers:", len(student_stats['duplicates']))
    print()
    print("Not assigned to BML SLBer:")
    for i in sorted(student_stats["other_students"]):
        print(i, student_data[i])
    print()
    print("More than 1 SLBer:")
    for i in sorted(student_stats["duplicates"]):
        print(i, student_data[i])
    print()
    print("More than 1 BML-R SLBer:")
    for student in sorted(student_stats["duplicates"]):
        slbers = student_data[student]['slb']
        res = [i for i in slbers if i in slb_list]
        if len(res) > 1:
            print(student_data[student])
    print()
    print("*" * 40)
    print()


def get_student_stats(student_data, slb_list):
    student_stats = {}
    all_students = [i for i in student_data]
    bml_students = []
    for student_num in student_data:
        slbers = student_data[student_num]['slb']
        for slber in slbers:
            if slber in slb_list:
                bml_students.append(student_num)
                break
    other_students = [i for i in student_data if i not in bml_students]
    duplicates = [i for i in student_data if len(student_data[i]['slb']) > 1]
    to_write = bml_students + other_students
    student_stats['all_students'] = all_students
    student_stats['bml_students'] = bml_students
    student_stats['other_students'] = other_students
    student_stats['duplicates'] = duplicates
    student_stats['to_write'] = to_write
    return student_stats


def write_excel(student_data, student_stats, slb_list, outfile):
    workbook = xlsxwriter.Workbook(outfile + ".xlsx")
    worksheet = workbook.add_worksheet()
    col = 0
    row = 1
    for student in sorted(student_stats['to_write']):
        col = 0
        slb = [i for i in student_data[student]["slb"] if i in slb_list]
        if slb:
            slb = ", ".join([i for i in student_data[student]["slb"] if i in slb_list])
        else:
            slb = ", ".join(student_data[student]["slb"])
        row_data = [student, student_data[student]['last_name'], student_data[student]["first_name"], slb]
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
    print("currently parsing:", in_file)
    out_file = args.output_file
    slb_list = read_slb(slb_file)
    pdf_content = extract_pdf(in_file)
    student_data = parse_text(pdf_content)
    student_stats = get_student_stats(student_data, slb_list)
    if verbose_status:
        print_stats(student_data, student_stats, slb_list)
    write_excel(student_data, student_stats, slb_list, out_file)
    print()
    print("Data written to", out_file + ".xlsx")
    print("Done")
    print("*" * 20)


if __name__ == "__main__":
    main()
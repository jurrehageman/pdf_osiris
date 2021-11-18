import pdfplumber
import argparse
import xlsxwriter

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
    args = parser.parse_args()
    return args


def extract_pdf(infile):
    data = []
    with pdfplumber.open(infile) as pdf:
        for num, page in enumerate(pdf.pages):
            if num > 0:
                text = page.extract_text()
                data.append(text)
    return data


def parse_text(data):
    student_slb = {}
    for num, line in enumerate(data):
        item = line.split("\n")
        slb = item[2][-5: -1]
        students = item[6:-1]
        if students:
            for student in students:
                student_num = student[0:6].strip()
                comma_pos = student.find(",")
                last_name = student[7:comma_pos].strip()
                BO_pos = student.find(" BO ")
                first_name = student[comma_pos + 1: BO_pos].strip()
                student_slb[student_num] = (slb, last_name, first_name)
    return student_slb


def write_excel(data, outfile):
    workbook = xlsxwriter.Workbook(outfile + ".xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.autofilter('A1:D1')
    header = ["Student Nummer", "Achternaam", "Voornaam", "SLB-er"]
    row = 0
    col = 0
    for i in header:
        worksheet.write(row, col, i, bold)
        col += 1
    row = 1
    for student in sorted(data):
        #print("processing:", student, data[student][1], data[student][2], data[student][0])
        col = 0
        row_data = [student, data[student][1], data[student][2], data[student][0]]
        for i in row_data:
            worksheet.write(row, col, i)
            col += 1
        row += 1
    workbook.close()


def main():
    args = get_comm_args()
    in_file = args.pdf_file
    out_file = args.output_file
    pdf_content = extract_pdf(in_file)
    student_data = parse_text(pdf_content)
    write_excel(student_data, out_file)
    print()
    print("Data written to", out_file + ".xlsx")
    print("Done")
    print("*" * 20)


if __name__ == "__main__":
    main()
import pdfplumber
import datetime
import xlsxwriter

sections = ('objective', 'skills', 'education', 'work', 'projects', 'achievements', 'certifications', 'coursework', 'responsibility')

def doc_parser():


    # pdf_content = ""
    words = list()
    with pdfplumber.open("Resume_Devershiprakash.pdf") as pdf:
        pages = pdf.pages
        for page in pages:
            # pdf_content += page.extract_text()
            for e in page.extract_words():
                words.append(e['text'])

    index = dict()
    # words = pdf_content.split()
    # print(len(words))
    for section in sections:
        for word in words:
            if section.casefold() == word.casefold() or section.casefold()+':' == word.casefold():
                index[section] = words.index(word)

    index_list = sorted(index.items(), key=lambda x: x[1])
    index = dict(index_list)
    print(words)
    result = dict()
    headings = list(index.keys())
    list_index = list(index.values())
    l = len(headings)-1
    for i in range(l):
        content = ' '.join(words[list_index[i]+1: list_index[i+1]])
        result[headings[i]] = content

    result[headings[-1]] = ' '.join(words[list_index[-1]+1: ])

    links = list()
    for word in words:
        if word.__contains__('@') or word.__contains__('.com') or word.__contains__('http'):
            links.append(word)

    for k in result:
        print(k, result[k], sep=':')

    return result, links


def resume_reader():

    result, links = doc_parser()
    print("CHOOSE:")
    print("0 : Links")
    for i in range(len(sections)):
        print(i+1, ':', sections[i])
    print("-1: To EXIT!")

    ct = datetime.datetime.now().timestamp()

    # Create an new Excel file and add a worksheet.
    session_filename = "resume_"+str(ct)+".xlsx"
    workbook = xlsxwriter.Workbook(session_filename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 200)

    count = 1
    while True:
        try:
            choice = int(input("Press any of the given number: "))

            if choice < 0:
                print("EXITING!")
                break
            if choice > len(sections):
                print("Please Enter a Valid Choice")
                continue

            key = ''
            if choice == 0:
                key = 'links'
                output = links
            else:
                key = sections[choice-1]
                output = result.get(key, "No Results Found")

            # Write text to Excel File
            cell_k = 'A'+str(count)
            cell_d = 'B'+str(count)
            worksheet.write(cell_k, key)
            worksheet.write(cell_d, output)

            print(key, ":", output)
            count+=1
        except:
            print("Only Numbers Please!!")

    workbook.close()


resume_reader()

import requests, bs4, openpyxl, sys #pyperclip?

while True:
    print('Name of file: ', end='')
    file_name = input()
    file_name_ext = file_name + '.xlsx'
    while True:
        sheet_index = 0
        try:
            wb = openpyxl.Workbook()
            wb.save(file_name_ext)
            print('Enter website to scrape: ', end='')
            site = input()

            res = requests.get(site)
            res.raise_for_status()

            data = bs4.BeautifulSoup(res.text, 'html.parser')

            elems = data.select('td')
            
            print('Name of worksheet: ', end='')
            sheet_name = input()
            wb = openpyxl.load_workbook(file_name_ext)
    #        wb = openpyxl.load_workbook('example2.xlsx')
            wb.create_sheet(index=sheet_index, title=sheet_name)

            sheet_index += 1
            sheet = wb[sheet_name]

            rowNum = 3
            counter = 1
            print('Regular or SOS (0 or 1): ', end='')
            reg_or_sos = int(input())
            if reg_or_sos == 0:
                for j in range(1, len(elems), 8):
                    print(elems[j].getText() + ': ' + (elems[j + 1].getText()))
                    print(counter)
                    sheet.cell(row=rowNum, column=1).value = elems[j].getText()
                    sheet.cell(row=rowNum, column=2).value = float(elems[j + 1].attrs['data-sort'])
                # sheet.cell(row=rowNum, column=2).value = float(elems[j + 1].getText())
                    rowNum += 1
                    counter += 1
                    if counter == 354:
                        break
            else:
                #SOS loop
                for j in range(1, len(elems), 6):
                    print(elems[j].getText() + ': ' + (elems[j + 1].getText()))
                    print(counter)
                    sheet.cell(row=rowNum, column=1).value = elems[j].getText()
                    #sheet.cell(row=rowNum, column=2).value = float(elems[j + 1].attrs['data-sort'])
                    sheet.cell(row=rowNum, column=2).value = float(elems[j + 1].getText())
                    rowNum += 1
                    counter += 1
                    if counter == 354:
                        break
            
            print('Enter season years: ', end='')
            season_years = input()
            sheet.cell(row=1, column=1).value = season_years
            sheet.cell(row=2, column=1).value = 'Team'
            sheet.cell(row=2, column=2).value = sheet_name
            
            print('Save file name: ', end='')
            save_file_name = input()
            save_file_name += '.xlsx'
            wb.save(save_file_name)
        except KeyboardInterrupt:
            print('NEW WEBSITE')
            sys.exit()



#!/usr/bin/env python3
import pandas as pd


def main(week):
    
    members = []
    # Spreadsheet from Google Drive
    member_file = 'Breaking Bread (Responses).xlsx'
    df = pd.read_excel(member_file, sheet_name='Form Responses 1')
    for i, name in enumerate(df[df.columns[1]]): # For each row
        # Columns: Timestamp, Name, Email, Phone #, Comments, Gender, Out This Week
        members.append((df[df.columns[1]][i], df[df.columns[5]][i], df[df.columns[6]][i], df[df.columns[2]][i], df[df.columns[3]][i]))
    
    males = [a for a in members if a[1] == 'M']
    females = [a for a in members if a[1] == 'F']
    
    tokens = []
    with open('pairs.txt', 'r') as pair_file:
        for line in pair_file:
            if line[:-1].split(', ')[0] == week:
                token = line[:-1].split(', ')[1:]
                # print(token)
                name1 = [(a, b, d, e) for (a, b, c, d, e) in males if a == token[0]]
                name2 = [(a, b, d, e) for (a, b, c, d, e) in females if a == token[0]]
                name3 = [(a, b, d, e) for (a, b, c, d, e) in males if a == token[1]]
                name4 = [(a, b, d, e) for (a, b, c, d, e) in females if a == token[1]]
                # print(name1, name2, name3, name4)
                if not name1:
                    name1 = name2
                if not name3:
                    name3 = name4
                tokens.append((name1, name3))
    
    with open('email.txt', 'r+') as email_file:
        for token in tokens:
            token = (token[0][0], token[1][0])
            # print(token)
            len_name_1 = len(token[0][0])
            len_email_1 = len(token[0][2])
            len_phone_1 = len(str(int(token[0][3])))
            len_name_2 = len(token[1][0])
            len_email_2 = len(token[1][2])
            len_phone_2 = len(str(int(token[1][3])))
            if len_name_1 >= len_email_1 and len_name_1 >= len_phone_1: # Name is larger
                name_line_1 = '| %s ' % token[0][0]
                address_line_1 = '| %s%s ' % (token[0][2], ' '*(len_name_1 - len_email_1))
                phone_line_1 = '| %d%s ' % (int(token[0][3]), ' '*(len_name_1 - len_phone_1))
            elif len_email_1 >= len_name_1 and len_email_1 >= len_phone_1: # Email is larger
                name_line_1 = '| %s%s ' % (token[0][0], ' '*(len_email_1 - len_name_1))
                address_line_1 = '| %s ' % token[0][2]
                phone_line_1 = '| %d%s ' % (int(token[0][3]), ' '*(len_email_1 - len_phone_1))
            else:
                name_line_1 = '| %s%s ' % (token[0][0], ' '*(len_phone_1 - len_name_1))
                address_line_1 = '| %s%s ' % (token[0][2], ' '*(len_phone_1 - len_email_1))
                phone_line_1 = '| %d ' % int(token[0][3])
            
            if len_name_2 >= len_email_2 and len_name_2 >= len_phone_2: # Name is larger
                name_line_2 = '| %s |' % token[1][0]
                address_line_2 = '| %s%s |' % (token[1][2], ' '*(len_name_2 - len_email_2))
                phone_line_2 = '| %d%s |' % (int(token[1][3]), ' '*(len_name_2 - len_phone_2))
            elif len_email_2 >= len_name_2 and len_email_2 >= len_phone_2: # Email is larger
                name_line_2 = '| %s%s |' % (token[1][0], ' '*(len_email_2 - len_name_2))
                address_line_2 = '| %s |' % token[1][2]
                phone_line_2 = '| %d%s |' % (int(token[1][3]), ' '*(len_email_2 - len_phone_2))
            else:
                name_line_2 = '| %s%s |' % (token[1][0], ' '*(len_phone_2 - len_name_2))
                address_line_2 = '| %s%s |' % (token[1][2], ' '*(len_phone_2 - len_email_2))
                phone_line_2 = '| %d |' % int(token[1][3])
            name_line = name_line_1 + name_line_2
            address_line = address_line_1 + address_line_2
            phone_line = phone_line_1 + phone_line_2
            print(name_line)
            print(address_line)
            print(phone_line)
            fill_line = '-'*len(name_line)
            print(fill_line)
            email_file.write('Hi!\n\nHere is your assigned breaking bread partner for this week:\n\n')
            email_file.write('%s\n%s\n%s\n%s\n%s\n%s\n%s\n\n' % (fill_line, name_line, fill_line, address_line, fill_line, phone_line, fill_line))
            email_file.write('If you wish to opt out of breaking bread (temporarily or permanently), or if you would like to participate for a single week at a time, contact Alex Diaz at (redacted).\n\n')
            

if __name__ == '__main__':
    main('6')

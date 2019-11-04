#!/usr/bin/env python3
import itertools
import math
import pandas as pd
import random


def week(num):
    members = []
    # Spreadsheet from Google Drive
    member_file = 'Breaking Bread (Responses).xlsx'
    df = pd.read_excel(member_file, sheet_name='Form Responses 1')
    for i, name in enumerate(df[df.columns[1]]): # For each row
        # Columns: Timestamp, Name, Email, Phone #, Comments, Gender, Out This Week
        members.append((df[df.columns[1]][i], df[df.columns[5]][i], df[df.columns[6]][i], df[df.columns[2]][i], df[df.columns[3]][i]))
    
    males = [a for a in members if a[1] == 'M' and a[2] != 'yes']
    females = [a for a in members if a[1] == 'F' and a[2] != 'yes']
    
    pairs = []
    # To remember what pairs have been made
    with open('pairs.txt', 'r') as used_pairs:
        for line in used_pairs:
            attributes = line.split(',')
            pairs.append((attributes[1][1:].strip('\n'), attributes[2][1:].strip('\n')))
    
    # print('Members', members)
    # print('Males', males)
    # print('Females', females)
    # print('Used Pairs', pairs)
    # print('%d Males/%d Females' % (len(males), len(females)))

    # Find every pairing of guys and girls
    male_combs = list(itertools.combinations(males, 2))
    random.shuffle(male_combs)
    female_combs = list(itertools.combinations(females, 2))
    random.shuffle(female_combs)
    
    # print('Male Combinations', male_combs)
    # print('Female Combinations', female_combs)
    
    new_pairs = []
    used_members = []
    
    # Create pairings for this week
    for possibility in male_combs:
        if (possibility[0][0], possibility[1][0]) not in pairs and possibility[0][0] not in used_members and possibility[1][0] not in used_members:
            new_pairs.append((possibility[0][0], possibility[1][0]))
            used_members.append(possibility[0][0])
            used_members.append(possibility[1][0])
    
    for possibility in female_combs:
        if (possibility[0][0], possibility[1][0]) not in pairs and possibility[0][0] not in used_members and possibility[1][0] not in used_members:
            new_pairs.append((possibility[0][0], possibility[1][0]))
            used_members.append(possibility[0][0])
            used_members.append(possibility[1][0])
    
    # print(used_members)
    print("# Males:", len(males), "\n# Females:", len(females))
    print(new_pairs)
    
    # Write to pairs.txt
    with open('pairs.txt', 'a') as pair_file:
        for pair in new_pairs:
            pair_file.write('%d, %s, %s\n' % (num, pair[0], pair[1]))


if __name__ == '__main__':
    week(6)

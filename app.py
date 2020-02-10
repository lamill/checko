import checko as ck

check_list = []
with open('list.txt', 'r') as f:
    for line in f:
        line=line.replace('\n','')
        check_list.append(line)

ck.make_xls(check_list)

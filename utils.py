def period_index(s: str):
    return int(s[1:])

def col_num_to_letter(col_num):
    letter = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def ensure(condition, message):
    if condition == False:
        print(f'\033[91m⚠ {message}\033[0m')
        exit()
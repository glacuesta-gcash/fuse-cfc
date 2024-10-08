import asyncio
import re

def is_period(n: str) -> bool:
    pattern = r'^p\d+$'
    return bool(re.match(pattern, n))

def period_index(s: str):
    """1-based, so p1 will return 1."""
    return int(s[1:])

def col_num_to_letter(col_num):
    letter = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def row_col_to_cell_ref(row, col) -> str:
    return f'{col_num_to_letter(col)}{row}'

def ensure(condition, message):
    """Will fail if condition is false."""
    if condition == False:
        print(f'\033[91m⚠ {message}\033[0m')
        exit()

async def parallel_calls(*partials):
    coroutines = [
        asyncio.create_task(
            asyncio.to_thread(partial)
        )
        for partial in partials
    ]
    tasks = await asyncio.gather(*coroutines)
    return tasks
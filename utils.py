def format_question_number(n: int, part: int) -> str:
    """Formats a zero-indexed question number"""
    if part == 0:
        # top left question number: 1, 2, 3 etc.
        return str(n+1)
    elif part == 1:
        # part of question: a, b, c, d etc.
        return 'abcdefghijklmnopqrstuvwxyz'[n%26]
    elif part == 2:
        # sub-part in roman numerals: i, ii, iii, iv, v etc.
        return ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'xi', 'xii', 'xiii', 'xiv', 'xv'][n%15]
    else:
        # back to numbers
        return str(n+1)

Q_COLUMNS = {
    0: 'A',
    1: 'B',
    2: 'C',
    3: 'D'
}
COLUMNS = {
    'total marks': 'F',
    'mark attained': 'E',
    'topic': 'J',
    'type': 'I',
    'percent': 'G',
    'time': 'H',
    'given answer': 'K'
}
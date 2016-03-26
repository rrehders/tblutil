import re

postalregex = re.compile(r'''(?!.*[DFIOQU])      # Eliminate invalid starting letters
    ([A-VXY]             # Valid starting letters
    \d                  # Number
    [A-Z])              # Letter
    .?                  # Optional seperator
    (\d                 # Number
    [A-Z]               # Letter
    \d                  # Number
    )
    ''', re.IGNORECASE | re.VERBOSE)

test = 'M6S 1J4'
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = ' M6S 1J4'
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = 'M6S 1J4 '
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = 'M6S1J4'
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = 'M6S1J4 '
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = ' M6S1J4'
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

test = ' M6S1J4 '
result = postalregex.sub(r'\1 \2', test.strip())
print('\'' + test + '\' \'' + result + '\'')

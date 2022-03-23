import os
path = os.getcwd()
print(path)

path = 'example.xlsx'
separate = os.path.splitext(path)
print(separate)
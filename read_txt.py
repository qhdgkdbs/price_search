f = open("test.txt", "r")
lines = f.read().split('\n')

for n,i in enumerate(lines):
    lines[n] = i.replace('\n','')
    if "hello" in lines[n]:
        print(i)

# print(i)


# print(lines)
#f = open("mail_user.txt", "a")
#f.write("hello"+",")
#f.close()

f = open("mail_user.txt", "r")
x=f.read()
y=str(x)
string = y[:len(x) - 1]
list_1 = string.split(",")
print("USERS :",list_1)
ste = list_1[:-1]
print("minus",ste)

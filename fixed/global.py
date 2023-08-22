def s():
  global r
  print("this is :",r)

def myd():
  global r
  r="rabbie"
  

def myfunc():
  global r
  r = "fantastic"

x = input("Enter your name:")
print("Hello, " + x)
if x=="a":
  myd()
  s()
elif x=="b":
  myfunc()
  s()
else:
  print("surya")

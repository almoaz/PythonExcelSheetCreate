import pyrebase
#auth.uid != null

const  = {
  "apiKey": "AIzaSyBLOfT0N-JTkm79S-XKvTIMvw3gNBBLsis",
  "authDomain": "smuct-6c3ad.firebaseapp.com",
  "databaseURL": "https://smuct-6c3ad-default-rtdb.firebaseio.com/",
  "projectId": "smuct-6c3ad",
  "storageBucket": "smuct-6c3ad.appspot.com",
  "messagingSenderId": "1013082282201",
  "appId": "1:1013082282201:web:90172c1a8806071c23c6c8",
  "measurementId": "G-BP3NWNMGT1"};

firebase = pyrebase.initialize_app(const)
db = firebase.database()

'''auth = firebase.auth()

email = input("write an email : ")
password = input("write an password : ")

try:
  auth.sign_in_with_email_and_password(email,password)
  print("Log in successfully!")
except:
  print("Invalid user or password. try agein")'''


teacher = 'ABCDEFABCDEFABCDEFABCDEF'
for x in range(18):
  t = teacher[x]

  if x<6:
    data = db.child("Department Of CSE And CSIT").child("Routine").child("Saturday").child(t+"1C").get()
    print(t+"1C: "+data.val())

  if x>5 and x<12:
    data = db.child("Department Of CSE And CSIT").child("Routine").child("Saturday").child(t+"1T").get()
    print(t+"1T: "+data.val())

  if x>11 and x<18:
    data = db.child("Department Of CSE And CSIT").child("Routine").child("Saturday").child(t+"1R").get()
    print(t+"1R: "+data.val())





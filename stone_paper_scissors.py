import random
import win32com.client
import streamlit as st
speaker=win32com.client.Dispatch("SAPI.SpVoice")

d=0
b=0
while(True):
 st.write("Enter 0 to exit")
 st.write("Enter 1 to choose stone\nEnter 2 to choose paper\nEnter 3 to choose scissors")
 a=int(st.text_input("Enter choice"))
 if(a>3 or a<0):
     raise ValueError("Invalid Input")
 else:
 
 
  c=random.randint(1,3)
  match a:
    case 1:
        if(c==1):
            st.write("Both choose stone.Tied")
            speaker.Speak("Both choose stone Tied")
        elif(c==2):
            st.write("Computer choose paper.You lost")
            speaker.Speak("Computer choose paper You lost")
        elif(c==3):
            d+=1
            st.write("Computer choose scissors.You Won")
            speaker.Speak("Computer choose scissors You Won")
    case 2:
        if(c==1):
            d+=1
            st.write("Computer choose stone.You Won")
            speaker.Speak("Computer choose stone You Won")
        elif(c==2):
            st.write("Both choose Paper.Tied")
            speaker.Speak("Both choose Paper Tied")
        elif(c==3):
            st.write("Computer choose scissors.You Lost")
            speaker.Speak("Computer choose scissors You lost")
    case 3:
        if(c==1):
            st.write("Computer choose stone.You Lost")
            speaker.Speak("Computer choose stone You lost")
        elif(c==2):
            d+=1
            st.write("Computer choose paper.You Won")
            speaker.Speak("Computer choose paper You Won")
        elif(c==3):
            st.write("Both choose Scissors.Tied")  
            speaker.Speak("Both choose Scissors Tied")
 b+=1                  
 if(a==0):
     break     
st.write(f"You have played {b} times") 
speaker.Speak(f"You have played {b} times")     
st.write(f"You have won {d} times")
speaker.Speak(f"You have won {d} times")

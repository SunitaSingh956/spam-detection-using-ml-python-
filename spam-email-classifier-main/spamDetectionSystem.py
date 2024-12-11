import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
import pythoncom
from win32com.client import Dispatch

def add_bg_from_url():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-size: cover
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

add_bg_from_url()

def speak(text):
    # Initialize COM in a thread-safe manner
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    
    try:
        # Create the speech object
        speak = Dispatch("SAPI.SpVoice")
        speak.Speak(text)
    finally:
        # Uninitialize COM after the operation
        pythoncom.CoUninitialize()

# Load the model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

def main():
    st.title("Email Spam Detection Application")
    activities = ["Classification", "About"]
    choices = st.sidebar.selectbox("Select Activities", activities)
    
    if choices == "Classification":
        st.subheader("Classification")
        msg = st.text_area("Enter text here")
        
        if st.button("Predict"):
            print(msg)
            print(type(msg))
            data = [msg]
            print(data)
            
            # Transform the input message using the vectorizer
            vec = cv.transform(data).toarray()
            
            # Predict the result using the model
            result = model.predict(vec)
            
            if result[0] == 0:
                st.success("This is Not A Spam Email")
                speak("This is Not A Spam Email")
            else:
                st.error("This is A Spam Email")
                speak("This is A Spam Email")
    
    if choices == "About":
        st.subheader("About")
        st.write("This is a spam email classifier application.")
        st.write("Made with Streamlit")
        st.write("By Sunita Singh")

if __name__ == "__main__":
    main()

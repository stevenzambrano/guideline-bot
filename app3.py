import streamlit as st
import tempfile
import docx 
import os
from docx import Document
from win32com import client as win32
import re

import nltk
nltk.download('punkt_tab')


def para_dic(temp_file_path):

    para_dic = {}
    doc = docx.Document(temp_file_path)

    # iterate over the index and paragraphs
    for index,paragraph in enumerate(doc.paragraphs):
        paragraph_text = paragraph.text.strip() # split the data into paragraphs and remove whitespaces

    #use nltk to tokenize your paragraphs into sentences!

        sentences = nltk.sent_tokenize(paragraph_text)

        #initialise a sentence dic
        sentence_dic = {}

    # then  you need to remember to capture the index of the paragraph, sentence!
        for i, sentence in enumerate(sentences):
            # Remove special bullet-like characters (e.g., )
            sentence = re.sub(r'[^\x00-\x7F]+', '', sentence)  # Remove non-ASCII characters
            # Remove numeric bullet points (e.g., 3.)
            sentence = re.sub(r'^\d+\.\s*', '', sentence)  # Remove number bullet points at the start
            # Remove extra spaces
            sentence = re.sub(r'\s+', ' ', sentence).strip()  # Clean up spaces
            # clean_sentence = re.sub(r'\s+', ' ', sentence).strip() # use regex to remove extra white spaces be

            # check if the sentence is not empty and contains more than just numbers (becareful as restrictions  may come in tables with numbers)
            if bool(sentence) and not re.match(r'^\d+\s*$', sentence):    
                # print(index, i, sentence, bool(sentence))
                sentence_dic[i] = sentence
        para_dic[index]= sentence_dic

    para_dic =  {k:v for k,v in para_dic.items() if v}
    return para_dic


def restriction_para(para_dic):
    restriction_dict = {}
    for keys,values in para_dic.items():
        temp_dict = {}
        for k,v in values.items():
            # print(keys, k,v)
            if "PV01" in v: #here we will ask the bot if it's an investment restriction and store in the temp dictionary
                print(keys,k,v)
                temp_dict[k]=v
        restriction_dict[keys]=temp_dict  # then we will store all the sentences which the llm believes is an investment rest to embed.

    restriction_dict={keys:values for keys,values in restriction_dict.items() if values}
    
    return restriction_dict


def annotate_document(temp_file_path,save_path,restriction_dict):

    word_app = win32.Dispatch("Word.Application")
    word_app.Visible=False
    doc3 = word_app.Documents.Open(temp_file_path)

    # Label and progress bar for annotation
    total_paragraphs = len(restriction_dict)
    progress_placeholder = st.empty()

    with progress_placeholder.container():
        st.write("Step 3: Annotating document with comments based on rule IDs")
        annotation_progress_bar = st.progress(0)  # Initialize the progress bar

    for i, (para_index, sentences) in enumerate(restriction_dict.items()):
        para = doc3.Paragraphs(para_index+1) # retrieves the line I need to annotate
        start= para.Range.Start
        end = para.Range.End
        sen_range = doc3.Range(start,end)
        doc3.Comments.Add(sen_range,"hi")

        # Update the progress bar
        progress = (i + 1) / total_paragraphs
        annotation_progress_bar.progress(progress)


    doc3.SaveAs2(save_path) # Change to your desired output path
    doc3.Close()
    word_app.Quit()

    # Complete the progress bar
    annotation_progress_bar.progress(1)
    progress_placeholder.empty()

    

def main():
 
    st.set_page_config(page_title="Guideline Bot", page_icon="🤖", layout="wide")

    st.title("Guideline bot")
    st.caption("by Steven Zambrano 🚀")
    bot_image_path = r"images/bot.png"
    st.image(bot_image_path,use_column_width=True)

    #chat interface
    uploaded_doc = st.file_uploader("Upload the Word Document", type=["docx"])


    if uploaded_doc:
        st.write("Analyzing document and adding comments")

        if st.button("Annotate Document"):
              try:
                # create a temporary file to save the uploaded document
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                    temp_file.write(uploaded_doc.getbuffer())
                    temp_file_path = temp_file.name # path to the temp file

           
                save_path = os.path.join(tempfile.gettempdir(), f"annotated_{os.path.basename(uploaded_doc.name)}")
                
                #call the paragraph dictionary
                para_dictionary = para_dic(temp_file_path)

                #call the restriction dictionary
                restrict_dic = restriction_para(para_dictionary)
                
                #here is the function we need
                annotate_document(temp_file_path,save_path, restrict_dic)


                with open(save_path,"rb") as file:
                    st.download_button(
                        label = "Download Annotated Document",
                        data = file,
                        file_name=save_path,
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
              except Exception as e:
                  st.error(f"error {e}")
                  
        
                
if __name__ == '__main__':
    main()




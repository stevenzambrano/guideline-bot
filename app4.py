import streamlit as st
import tempfile
import os
from win32com import client as win32
import re
from LLM import classify_sentence, get_embedding
from mypine import rule_ids
import nltk

nltk.download('punkt_tab')

def para_dic(temp_file_path):
    # Create a dictionary to store para_dic = {para_index: {sentence_index: text}}
    para_dic = {}

    # Open the document with win32com
    word_app = win32.Dispatch("Word.Application")
    doc = word_app.Documents.Open(temp_file_path)

    # Iterate over the paragraphs
    for index, paragraph in enumerate(doc.Paragraphs):
        paragraph_text = paragraph.Range.Text.strip()  # Get paragraph text

        # Check if paragraph contains tables
        if paragraph.Range.Tables.Count > 0:
            continue  # Skip paragraphs containing tables

        # Tokenize the paragraph into sentences
        sentences = nltk.sent_tokenize(paragraph_text)
        sentence_dic = {}

        for i, sentence in enumerate(sentences):
            # Clean the sentence text
            sentence = re.sub(r'[^\x00-\x7F]+', '', sentence)  # Remove non-ASCII characters
            sentence = re.sub(r'^\d+\.\s*', '', sentence)  # Remove number bullet points at the start
            sentence = re.sub(r'\s+', ' ', sentence).strip()  # Clean up spaces

            # Check if the sentence is not empty and contains more than just numbers
            if sentence and not re.match(r'^\d+\s*$', sentence):
                sentence_dic[i] = sentence

        # Add the sentence dictionary to para_dic
        if sentence_dic:
            para_dic[index] = sentence_dic

    doc.Close()  # Close the document
    word_app.Quit()  # Quit Word application

    return para_dic


def restriction_para(para_dic):
    res = {}

    total_sentences = sum(len(v) for v in para_dic.values())
    processed_sentences = 0
    progress_placeholder = st.empty()  # Single placeholder for both label and progress bar

    with progress_placeholder.container():
        st.write("Classifying sentences for investment restrictions")
        classification_progress_bar = st.progress(0)

    for i, v in para_dic.items():
        temp2 = {}
        for ide, sen in v.items():
            result = classify_sentence(sen)  # Assume this function is defined
            if result == "investment restriction":
                temp2[ide] = sen
            processed_sentences += 1
            progress = processed_sentences / total_sentences
            classification_progress_bar.progress(progress)

        res[i] = temp2

    progress_placeholder.empty()  # Clear label and progress bar after completion

    # Clean up the restriction dictionary
    return {ke: va for ke, va in res.items() if va}


def vectorize_dict(restriction_dict):
    vectorized_dict = {}
    total_sentences = sum(len(v) for v in restriction_dict.values())
    processed_sentences = 0
    progress_placeholder = st.empty()  # Single placeholder for both label and progress bar

    with progress_placeholder.container():
        st.write("Generating vectors for classified sentences")
        vectorization_progress_bar = st.progress(0)

    for inde, valu in restriction_dict.items():
        temp3 = {}
        for index3, values3 in valu.items():
            vector = get_embedding(f'"{values3}"')  # Assume this function is defined
            temp3[index3] = vector
            processed_sentences += 1
            progress = processed_sentences / total_sentences
            vectorization_progress_bar.progress(progress)

        vectorized_dict[inde] = temp3

    progress_placeholder.empty()  # Clear label and progress bar after completion

    return vectorized_dict


def annotate_document(temp_file_path, save_path, rule_ids):
    word_app = win32.Dispatch("Word.Application")
    doc3 = word_app.Documents.Open(temp_file_path)

    # Initialize progress tracking
    total_paragraphs = len(rule_ids)
    progress_placeholder = st.empty()  # Single placeholder for both label and progress bar
    
    with progress_placeholder.container():
        st.write("Annotating document with comments based on rule IDs")
        annotation_progress_bar = st.progress(0)
   
    for i, (para_index, sentences) in enumerate(rule_ids.items()):
        para = doc3.Paragraphs(para_index + 1)  # Retrieves the paragraph to annotate
        start = para.Range.Start
        end = para.Range.End
        sen_range = doc3.Range(start, end)

        for sen_index, rules in sentences.items():
            doc3.Comments.Add(sen_range, f'Rule ID: {rules}')

        # Update the progress bar
        progress = (i + 1) / total_paragraphs
        annotation_progress_bar.progress(progress)

    doc3.SaveAs2(save_path)  # Change to your desired output path
    doc3.Close()
    word_app.Quit()
    progress_placeholder.empty()  # Clear label and progress bar after completion



def main():
    st.set_page_config(page_title="Guideline Bot", page_icon="🤖", layout="wide")

    st.title("Guideline bot")
    st.caption("by Steven Zambrano 🚀")
    bot_image_path = r"images/bot.png"
    st.image(bot_image_path, use_column_width=True)

    # Chat interface
    uploaded_doc = st.file_uploader("Upload the Word Document", type=["docx"])

    if uploaded_doc:
        st.write("Analyzing document and adding comments")

        if st.button("Annotate Document"):
            try:
                # Create a temporary file to save the uploaded document
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                    temp_file.write(uploaded_doc.getbuffer())
                    temp_file_path = temp_file.name  # Path to the temp file

                save_path = os.path.join(tempfile.gettempdir(), f"annotated_{os.path.basename(uploaded_doc.name)}")

                # Call the paragraph dictionary
                para_dictionary = para_dic(temp_file_path)

                # Call the restriction dictionary
                restrict_dic = restriction_para(para_dictionary)

                # Vectorize the restrictions
                vector_dict = vectorize_dict(restrict_dic)

                # Call pinecone to retrieve the rule ids
                rules_dict = rule_ids(vector_dict)

                # Call is the annotation functiond
                annotate_document(temp_file_path, save_path, rules_dict)

                with open(save_path, "rb") as file:
                    st.download_button(
                        label="Download Annotated Document",
                        data=file,
                        file_name=save_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Error: {e}")


if __name__ == '__main__':
    main()

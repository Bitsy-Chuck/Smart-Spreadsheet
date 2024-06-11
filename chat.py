import json

import openai as openai
import streamlit as st
import pandas as pd
from langchain.chat_models import ChatOpenAI
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
import os


from helper_functions import get_sheet_from_excel, get_table_ranges, process_simple_table, process_hierarchical_table


def initialize_session_state():
    """Initialize session state variables."""
    if "messages" not in st.session_state:
        st.session_state["messages"] = []
    if "conversation_memory" not in st.session_state:
        st.session_state["conversation_memory"] = ConversationBufferMemory(
            memory_key="chat_history"
        )


def display_chat_history():
    """Display the chat history from the session state."""
    for message in st.session_state["messages"]:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])


def load_data(file_path):
    """Load data from an Excel file."""
    return pd.read_excel(file_path)


def get_user_input():
    """Get user input from the chat interface."""
    return st.chat_input("Type your message...")


def main():
    """Main function to run the Streamlit app."""
    load_dotenv()
    api_key = os.getenv("ENV_VARIABLE")

    st.set_page_config(page_title="Excel Data Chatbot", page_icon=":bar_chart:")
    st.title("Excel Data Chatbot")

    initialize_session_state()

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    sheet_name = st.chat_input("sheet_name")
    if uploaded_file is not None:
        ws = get_sheet_from_excel(uploaded_file, sheet_name)
        ranges, s_wbs, c_wbs =get_table_ranges(ws)
        excel_table_data =[]
        for s_wb in s_wbs:
            data = process_simple_table(s_wb.active)
            excel_table_data.append(data)

        for c_wb in c_wbs:
            data = process_hierarchical_table(c_wb.active)
            excel_table_data.append(data)

        llm = ChatOpenAI(temperature=0.9, model="gpt-4", openai_api_key=api_key)

        chat(excel_table_data)

def chat(sysPrompt):
    import streamlit as st
    def response_generator(prompt):
        response = openai.Completion.create(
            engine="gpt-4o",
            prompt=prompt,
            max_tokens=50,
            n=1,
            stop=None,
            temperature=0.7,
        )
        response_text = response.choices[0].text.strip()
        return response_text

    st.title("Simple chat")

    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = []

    # Display chat messages from history on app rerun
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    # Accept user input
    if prompt := st.chat_input("What is up?"):
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": prompt})
        # Display user message in chat message container
        with st.chat_message("user"):
            st.markdown(prompt)

        # Display assistant response in chat message container
        prompt_to_agent = "This is the data on which you have to answer questions: "+ json.dumps(sysPrompt) + "\n User's prompt is : " + prompt
        with st.chat_message("assistant"):
            response = st.write_stream(response_generator(prompt_to_agent))
        # Add assistant response to chat history
        st.session_state.messages.append({"role": "assistant", "content": response})


if __name__ == "__main__":
    main()
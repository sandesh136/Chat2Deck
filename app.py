import streamlit as st
from src.interaction_service import query_to_pptx

def main():
    st.set_page_config(page_title="Chat2Deck", page_icon="🤖")
    st.title("Chat2Deck: AI Presentation Generator")
    
    query = st.text_area("Enter your query or topic:")
    num_slides = st.slider("Number of slides", 3, 10, 5)
    generate = st.button("Generate PPT")
    
    if generate:
        if not query.strip():
            st.warning("Please enter a non-empty query.")
            return
        
        with st.spinner("Generating PPT..."):
            try:
                ppt_file = query_to_pptx(query, num_slides)
                st.download_button("Download PPT", ppt_file, file_name="generated_presentation.pptx")
                st.success("Presentation generated successfully!")
            except RuntimeError as e:
                st.error(f"We are facing some technical problems. Please try again later.\n\nDetails: {e}")
        
if __name__ == "__main__":
    main()
